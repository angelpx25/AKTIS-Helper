<#
1. Retrieve additional ent and basic values in GetExistingAnalyticsData
2. Add function to determine if proposal exceeds subscriotion (ent and basic)
3. Post additional ent and basic values in PostToAktis 


#>

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

Function Show-Msgbox {
    Param(
        [string]$Message,
        [string]$Icon,
        [string]$Button,
        [string]$Title
    )
    
    #Buttons: OkOnly, OkCancel, AbortRetryIgnore, YesNoCancel, YesNo, RetryCancel
    #Icons: Critical, Question, Exclamation, Information

    [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
    [Microsoft.VisualBasic.Interaction]::Msgbox($Message,"$Icon,$Button",$Title)
}

function ThrowPostSuccessBalloon {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $ToastNotification = New-Object System.Windows.Forms.NotifyIcon
    $ToastNotification.Icon = [System.Drawing.SystemIcons]::Information
    $ToastNotification.BalloonTipText = "Successfully posted proposal statistics data for $Script:ProposalName to Aktis and Analytics Weekly Snapshot!"
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
    
    ExecuteAktisHelper
}    

function ExecuteAktisHelper {

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    
    $AktisHelperForm = New-Object System.Windows.Forms.Form
    $AktisHelperForm.Text = "Aktis Helper"
    $AktisHelperForm.Size = '1000,275'
    $AktisHelperForm.StartPosition = "CenterScreen"
    $AktisHelperForm.MinimumSize = $AktisHelperForm.Size
    $AktisHelperForm.MaximizeBox = $False
    $AktisHelperForm.Topmost = $True
    
    $DnDBoxLabel = New-Object Windows.Forms.Label
    $DnDBoxLabel.Location = '5,5'
    $DnDBoxLabel.AutoSize = $True
    $DnDBoxLabel.Text = "Drop Excel workplan here:"
    
    $DnDListBox = New-Object Windows.Forms.ListBox
    $DnDListBox.Location = '5,25'
    $DnDListBox.Height = 30
    $DnDListBox.Width = 972
    $DnDListBox.AllowDrop = $True
    
    $DnDExtractedDataBoxLabel = New-Object Windows.Forms.Label
    $DnDExtractedDataBoxLabel.Location = '5,65'
    $DnDExtractedDataBoxLabel.AutoSize = $True
    $DnDExtractedDataBoxLabel.Text = "Extracted data to be posted:"
    
    $DnDExtractedDataListBox = New-Object Windows.Forms.ListBox
    $DnDExtractedDataListBox.Location = '5,85'
    $DnDExtractedDataListBox.Height = 125
    $DnDExtractedDataListBox.Width = 972
    $DnDExtractedDataListBox.AllowDrop = $False
    
    $StatusBar = New-Object System.Windows.Forms.StatusBar
    $StatusBar.Text = "Ready"
    
    $AktisHelperForm.SuspendLayout()
    $AktisHelperForm.Controls.Add($DnDBoxLabel)
    $AktisHelperForm.Controls.Add($DnDListBox)
    $AktisHelperForm.Controls.Add($DnDExtractedDataBoxLabel)
    $AktisHelperForm.Controls.Add($DnDExtractedDataListBox)
    $AktisHelperForm.Controls.Add($StatusBar)
    $AktisHelperForm.ResumeLayout()
    
    $DnDListBox_DragOver = [System.Windows.Forms.DragEventHandler]{
        if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)){
            $_.Effect = 'Copy'
        }
        else{
            $_.Effect = 'None'
        }
    }
        
    $DnDListBox_DragDrop = [System.Windows.Forms.DragEventHandler]{
        ForEach ($FileName in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)){
            #$Script:ProposalPath = $FileName
            Get-FinancialDataFromWorkplan $FileName
            $DnDListBox.Items.Add("$Script:ProposalName --- $FileName")
            $DnDExtractedDataListBox.Items.Add("Proposal Name: $Script:ProposalName")
            $DnDExtractedDataListBox.Items.Add("Workplan location: $FileName")
            $DnDExtractedDataListBox.Items.Add("Proposed Hours: $Script:ProposedHours")
            $DnDExtractedDataListBox.Items.Add("Hourly Rate: $Script:HourlyRate")
            $DnDExtractedDataListBox.Items.Add("Total MRR: $Script:TotalMRR")
            $DnDExtractedDataListBox.Items.Add("PS Revenue: $Script:PSRevenue")
            $DnDExtractedDataListBox.Items.Add("Products Revenue: $Script:ProductsRevenue")
            $DnDExtractedDataListBox.Items.Add("Total NRR: $Script:TotalNRR")   
            $DnDExtractedDataListBox.Items.Add("--------------------------------------------------")
            $StatusBar.Text = ("List contains $($DnDListBox.Items.Count) item(s)")
            #MetadataTagPrompt
            LaunchPostConfirmationModal
        }
    }
    
    $AktisHelperForm_FormClosed = {
        try
        {
            $DnDListBox.remove_DragOver($DnDListBox_DragOver)
            $DnDListBox.remove_DragDrop($DnDListBox_DragDrop)
            #$AktisHelperForm.remove_FormClosed($AktisHelperForm_Cleanup_FormClosed)
        }
        catch [Exception]
        { }
    }
    
    $DnDListBox.Add_DragOver($DnDListBox_DragOver)
    $DnDListBox.Add_DragDrop($DnDListBox_DragDrop)
    $AktisHelperForm.Add_FormClosed($AktisHelperForm_FormClosed)

    $AktisHelperFormResult = $AktisHelperForm.ShowDialog()
    
        if($AktisHelperFormResult -eq [System.Windows.Forms.DialogResult]::OK){
            #OK action
        }
        elseif($AktisHelperFormResult -eq [System.Windows.Forms.DialogResult]::Cancel){
            $AktisHelperForm.Close()
            $AktisHelperForm.Dispose()
        }
        elseif(!($AktisHelperFormResult -eq [System.Windows.Forms.DialogResult]::OK)){
            $AktisHelperForm.Close()
        }
}

function Get-FinancialDataFromWorkplan ($ProposalFilePath) {

    try{

        if(!(Test-Path -Path "C:\Aktis")){
            New-Item -Path "C:\" -Name "Aktis" -ItemType "Directory" | Out-Null
        }
        if(!(Test-Path -Path "C:\Aktis\Temp")){
            New-Item -Path "C:\Aktis" -Name "Temp" -ItemType "Directory" | Out-Null
        }
        
        Copy-Item -Path $ProposalFilePath -Destination "C:\Aktis\Temp\$($ProposalFilePath | Split-Path -Leaf)" -Force

        $FilePath = Get-Item "C:\Aktis\Temp\$($ProposalFilePath | Split-Path -Leaf)"

        if($(Test-FileLock $FilePath) -eq "True"){
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following file is already open. Requesting to close file.`r`n$Filepath`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
            [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
            [void] [Microsoft.VisualBasic.Interaction]::MsgBox("$($FilePath | Split-Path -Leaf) is already open. Close the file and select OK.`r`n`r`nFull path:`r`n$Filepath", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')

            while(
                $(Test-FileLock $FilePath) -eq "True"
            ){
                Start-Sleep -Milliseconds "500"
            }
        }

        $ExcelOpen = New-Object -COMObject Excel.Application
        $ExcelOpen.Visible = $False
        $ExcelSpreadsheet = $ExcelOpen.Workbooks.Open($FilePath)

    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to open Excel file. Ensure Microsoft Excel is installed.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to open Excel file. Ensure file is closed. Ensure Microsoft Excel is installed.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()

    }

    try{
        $ProposalNameHeadingColumnLetter = (($ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Project Name:")).Address($False,$False)) -Replace "\d+"
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find the cell containing the following text:`r`n`r`nProject Name:`r`n`r`nError details:`r`n$($global:intErr++)Error:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to find the cell containing the following text:`r`n`r`nProject Name:`r`n`r`nError details:`r`n$($global:intErr++)Error:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()

    }

    $ProposalNameRow = $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Project Name:").Row + 1

    try{
        $Script:ProposalName = $ExcelSpreadsheet.Sheets.Item(1).Range("$ProposalNameHeadingColumnLetter" + "$ProposalNameRow").Value2
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to obtain the value of the Proposal Name cell.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to obtain the value of the Proposal Name cell.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
        
    }

    try{
        $Script:AktisKTNumber = ($ExcelSpreadsheet.Sheets.Item(1).Range("$ProposalNameHeadingColumnLetter" + "$ProposalNameRow").Comment.Shape.AlternativeText) -Replace "^.*?: "
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to obtain the value of the Proposal Name cell note, which is expected to contain the KT Number.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to obtain the value of the Proposal Name cell note, which is expected to contain the KT Number.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()

    }

    try{
        $QuantityColumnLetter = (($ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Quantity")).Address($False,$False)) -Replace "\d+"
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find the cell containing the following text:`r`n`r`nQuantity`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to find the cell containing the following text:`r`n`r`nQuantity`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
        
    }

    try{
        $PPUColumnLetter = (($ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("PPU")).Address($False,$False)) -Replace "\d+"
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find the cell containing the following text:`r`n`r`nPPU`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to find the cell containing the following text:`r`n`r`nPPU`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
        
    }

    try{
        $ExtendedPriceColumnLetter = (($ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Extended Price")).Address($False,$False)) -Replace "\d+"
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find the cell containing the following text:`r`n`r`nExtended Price`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to find the cell containing the following text:`r`n`r`nExtended Price`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
        
    }

    try{
        $NRRTotalRow = $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Non Recurring Fees Total:").Row
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find the cell containing the following text:`r`n`r`nNon Recurring Fees Total:`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to find the cell containing the following text:`r`n`r`nNon Recurring Fees Total:`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
        
    }

    try{
        $MRRTotalRow = $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Monthly Recurring Fees Total:").Row
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find the cell containing the following text:`r`n`r`nMonthly Recurring Fees Total:`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to find the cell containing the following text:`r`n`r`nMonthly Recurring Fees Total:`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
        
    }

    try{
        $ProposedHoursCellLocation = $QuantityColumnLetter + $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find(" Professional Services").Row
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find the cell containing the following text:`r`n`r`n Professional Services`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to find the cell containing the following text:`r`n`r`n Professional Services`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
        
    }

    $HourlyRateCellLocation = $PPUColumnLetter + $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find(" Professional Services").Row

    $TotalMRRCellLocation = $ExtendedPriceColumnLetter + $MRRTotalRow

    $PSRevenueCellLocation = $ExtendedPriceColumnLetter + $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find(" Professional Services").Row

    $TotalNRRCellLocation = $ExtendedPriceColumnLetter + $NRRTotalRow


    try{
        $Script:ProposalNumber = ($Script:ProposalName | Select-String "(?<=\[)[^]]+(?=\])").Matches.Value
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to extract proposal number from full proposal name.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to extract proposal number from full proposal name.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
        
    }
    
    try{
        $Script:ProposalVersion = $Script:ProposalName -Replace "^.*?] V"
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to extract proposal version from full proposal name.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to extract proposal version from full proposal name.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
        
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
    
    [decimal]$Script:ProductsRevenue = [math]::Round([int]$Script:TotalNRR - [int]$Script:PSRevenue,2)
   
    $ExcelSpreadsheet.Close($True)
    $ExcelOpen.Quit()

    Remove-Item -Path "C:\Aktis\Temp\$($ProposalFilePath | Split-Path -Leaf)"
    
}

function LaunchPostConfirmationModal {

    $AktisPostConfForm = New-Object System.Windows.Forms.Form
    $AktisPostConfForm.Text = "Aktis Helper - Post Confirmation"
    $AktisPostConfForm.Size = '500,250'
    $AktisPostConfForm.StartPosition = "CenterScreen"
    $AktisPostConfForm.MinimumSize = $AktisPostConfForm.Size
    $AktisPostConfForm.MaximizeBox = $False
    $AktisPostConfForm.Topmost = $True

    $PostConfBoxLabel = New-Object Windows.Forms.Label
    $PostConfBoxLabel.Location = '5,5'
    $PostConfBoxLabel.AutoSize = $True
    $PostConfBoxLabel.Text = "Do you want to post the following information to Aktis?"

    $PostConfListBox = New-Object Windows.Forms.ListBox
    $PostConfListBox.Location = '5,25'
    $PostConfListBox.Height = 110
    $PostConfListBox.Width = 472
    $PostConfListBox.AllowDrop = $True

    $PostButton = New-Object System.Windows.Forms.Button
    $PostButton.Location = '180,160'
    $PostButton.Size = '125,23'
    $PostButton.Width = 125
    $PostButton.Text = "Continue"

    $PostConfStatusBar = New-Object System.Windows.Forms.StatusBar
    $PostConfStatusBar.Text = "Ready to post..."

    $PostButton_Click = {
        GetExistingAnalyticsData 
    }

    $PostButton.Add_Click($PostButton_Click)

    $AktisPostConfForm.SuspendLayout()
    $AktisPostConfForm.Controls.Add($PostConfBoxLabel)
    $AktisPostConfForm.Controls.Add($PostConfListBox)
    $AktisPostConfForm.Controls.Add($PostButton)
    $AktisPostConfForm.Controls.Add($PostConfStatusBar)
    $AktisPostConfForm.ResumeLayout()

    #Format appropriate fields to use currency

    $PostConfListBox.Items.Add("Proposal Name: $Script:ProposalName")
    $PostConfListBox.Items.Add("Proposed Hours: $([decimal]$Script:ProposedHours)")
    $PostConfListBox.Items.Add("Hourly Rate: `$$([decimal]$Script:HourlyRate)")
    $PostConfListBox.Items.Add("Total MRR: `$$([decimal]$Script:TotalMRR)")
    $PostConfListBox.Items.Add("PS Revenue: `$$([decimal]$Script:PSRevenue)")
    $PostConfListBox.Items.Add("Products Revenue: `$$([decimal]$Script:ProductsRevenue)")
    $PostConfListBox.Items.Add("Total NRR: `$$([decimal]$Script:TotalNRR)")
    $PostConfListBox.Items.Add("--------------------------------------------------")

    $AktisPostConfFormResult = $AktisPostConfForm.ShowDialog()
        
        if($AktisPostConfFormResult -eq [System.Windows.Forms.DialogResult]::OK){
            #OK action
        }
        elseif($AktisPostConfFormResult -eq [System.Windows.Forms.DialogResult]::Cancel){
            $AktisPostConfForm.Close()
            $AktisPostConfForm.Dispose()
            $AktisHelperForm.Close()
            $AktisHelperForm.Dispose()
        }
        elseif(!($AktisPostConfFormResult -eq [System.Windows.Forms.DialogResult]::OK)){
            $AktisPostConfForm.Close()
            $AktisHelperForm.Close()
        }   
}

function GetExistingAnalyticsData {

    $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis..."

    try{
        $ProposalStatisticsRawData = (Get-PnPListItem -List "Proposal Statistics" -Fields "KTNumber","ID","ProposalName","MSP","Company","IntakeType","KTDate","DateProposalCompleted","DeliveredtoClient","Subscription","_x0035_Day","_x0033_DayExpedite","DateProposalWritten","DDELead","PresenttoClient_x002f_Staff","Draft1","Draft2","GratisRevision","ProposalExceedsSubscription_x002","ProposalExceedsSubscription_x0020").FieldValues
    
        $ProposalStatisticsData = @()
    
        $ProposalStatisticsRawData | ForEach-Object {
            $ProposalStatisticsData += $_
        }
    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
    }

    try{

        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(1%)"
        $Script:ProposalID = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).ID
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(7%)"
        $Script:LocalAnalyticsProposalName = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).ProposalName
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(14%)"
        $Script:ProposalMSP = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).MSP
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(21%)"
        $Script:ProposalMSPsClient = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).Company
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(28%)"
        $Script:ProposalIntakeType = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).IntakeType
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(35%)"
        $Script:ProposalKTDate = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).KTDate
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(42%)"
        $Script:ProposalDateCompleted = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).DateProposalCompleted
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(49%)"
        $Script:ProposalDeliveredToClient = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).DeliveredtoClient
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(56%)"
        $Script:ProposalSubscription = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).Subscription
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(63%)"
        $Script:ProposalBiOp5Day = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber})._x0035_Day
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(70%)"
        $Script:ProposalBiOp3DayExpedite = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber})._x0033_DayExpedite
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(75%)"
        $Script:DateProposalWritten = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).DateProposalWritten
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(79%)"
        $Script:ExistingDDELead = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).DDELead
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(87%)"
        $Script:ProposalBiOpPresentToClient = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).PresenttoClient_x002f_Staff
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(90%)"
        $Script:ProposalBiOpDraft1 = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).Draft1
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(94%)"
        $Script:ProposalBiOpDraft2 = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).Draft2
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(96%)"
        $Script:IsGratisRevision = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).GratisRevision
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(98%)"
        $Script:ProposalExceedsSubEnt = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).ProposalExceedsSubscription_x002
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(99%)"
        $Script:ProposalExceedsSubBasic = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:AktisKTNumber}).ProposalExceedsSubscription_x0020
        $PostConfStatusBar.Text = "Retrieving existing analytics data from Aktis...(100%)"

    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve existing analytics data for:`r`n$Script:ProposalName.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve existing analytics data for:`r`n$Script:ProposalName.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
    }

    $PostConfStatusBar.Text = "Successfully retrieved existing analytics data from Aktis!"
    RetrieveClockifyTimeData
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
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify project details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Clockify project details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
    }

    $ClockifyProjectDetails = @()
    $ClockifyProjectDetailsRawData | ForEach-Object {
      $ClockifyProjectDetails += $_
    }

    $PostConfStatusBar.Text = "Verifying Clockify project exists..."
    if(!($ClockifyProjectDetails.Name | Select-String "$Script:ProposalNumber V$Script:ProposalVersion")){
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Clockify project doesn't exist for $("$Script:ProposalNumber V$Script:ProposalVersion").`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Clockify project doesn't exist for $("$Script:ProposalNumber V$Script:ProposalVersion"). Please create it and select OK.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        RetrieveClockifyTimeData
    }
    
    $ClockifyProjectDetails | ForEach-Object {
        
        if($_.Name -eq "$Script:ProposalNumber V$Script:ProposalVersion"){
            $Script:ClockifyProjectID = $_.ID

            $PostConfStatusBar.Text = "Retrieving Clockify task level details..."
            GetProjectTaskLevelDetails

            $PostConfStatusBar.Text = "Retrieving Clockify user time data..."
            GetClockifyUserTimeEntryData

            $AngelRunningTimeEntryProjectID = ($Script:AngelTimeEntries | Where-Object {$Null -eq $_.TimeInterval.End}).ProjectID
            $AdamRunningTimeEntryProjectID = ($Script:AdamTimeEntries | Where-Object {$Null -eq $_.TimeInterval.End}).ProjectID
            $VictorRunningTimeEntryProjectID = ($Script:VictorTimeEntries | Where-Object {$Null -eq $_.TimeInterval.End}).ProjectID
    
            if($AngelRunningTimeEntryProjectID -eq $Script:ClockifyProjectID){
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "Angel currently has a running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n`r`nPrompting to stop currently running time entry.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Angel currently has running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n`r`n Please stop the running time entry and select OK.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                GetProjectTaskLevelDetails
                GetClockifyUserTimeEntryData
            }
    
            if($AdamRunningTimeEntryProjectID -eq $Script:ClockifyProjectID){
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "Adam currently has a running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n`r`nPrompting to stop currently running time entry.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Adam currently has running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n`r`n Please stop the running time entry and select OK.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                GetProjectTaskLevelDetails
                GetClockifyUserTimeEntryData
            }
    
            if($VictorRunningTimeEntryProjectID -eq $Script:ClockifyProjectID){
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "Victor currently has a running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n`r`nPrompting to stop currently running time entry.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Victor currently has running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n`r`n Please stop the running time entry and select OK.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                GetProjectTaskLevelDetails
                GetClockifyUserTimeEntryData
            }
        
            LaunchTimeEntryModal
    
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
        $Script:VictorUserID = (Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/users") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Body @{email = "varaoz@protonmail.com"}).ID      
    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify user IDs.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Clockify user IDs.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
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
          
          $VictorRawTimeEntryData = Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/user/" + $Script:VictorUserID + "/time-entries") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Body @{
            "start" = $DaysAgoISO8601
            "end" = $CurrentDateISO8601
            "page-size" = $PageSize
          }
    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify time entry data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Clockify time entry data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
    }

    $Script:AdamTimeEntries = @()
    $Script:AngelTimeEntries = @()
    $Script:VictorTimeEntries = @()
    
    $AdamRawTimeEntryData | ForEach-Object {
      $Script:AdamTimeEntries += $_
    }
    
    $AngelRawTimeEntryData | ForEach-Object {
      $Script:AngelTimeEntries += $_
    }
    
    $VictorRawTimeEntryData | ForEach-Object {
        $Script:VictorTimeEntries += $_
    }

    $AngelPCTime = ($Script:AngelTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPCTaskID}).TimeInterval.Duration
    $AngelPWTime = ($Script:AngelTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPWTaskID}).TimeInterval.Duration
    $AngelDCTime = ($Script:AngelTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyDCTaskID}).TimeInterval.Duration
    $AngelIPTime = ($Script:AngelTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyIPTaskID}).TimeInterval.Duration

    $AdamPCTime = ($Script:AdamTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPCTaskID}).TimeInterval.Duration
    $AdamPWTime = ($Script:AdamTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPWTaskID}).TimeInterval.Duration
    $AdamDCTime = ($Script:AdamTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyDCTaskID}).TimeInterval.Duration
    $AdamIPTime = ($Script:AdamTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyIPTaskID}).TimeInterval.Duration

    $VictorPCTime = ($Script:VictorTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPCTaskID}).TimeInterval.Duration
    $VictorPWTime = ($Script:VictorTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPWTaskID}).TimeInterval.Duration
    $VictorDCTime = ($Script:VictorTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyDCTaskID}).TimeInterval.Duration
    $VictorIPTime = ($Script:VictorTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyIPTaskID}).TimeInterval.Duration

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

    [decimal]$Script:VictorTotalPCTime = "0.00"
    [decimal]$Script:VictorTotalPWTime = "0.00"
    [decimal]$Script:VictorTotalDCTime = "0.00"
    [decimal]$Script:VictorTotalIPTime = "0.00"

    $VictorPCTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:VictorTotalPCTime = [decimal]$TimeSpent + [decimal]$Script:VictorTotalPCTime

    }

    $VictorPWTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:VictorTotalPWTime = [decimal]$TimeSpent + [decimal]$Script:VictorTotalPWTime

    }

    $VictorDCTime | ForEach-Object {

        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:VictorTotalDCTime = [decimal]$TimeSpent + [decimal]$Script:VictorTotalDCTime

    }

    $VictorIPTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:VictorTotalIPTime = [decimal]$TimeSpent + [decimal]$Script:VictorTotalIPTime

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

    [decimal]$VictorTotalTime = "0.00"
    $VictorTime = $Script:VictorTimeEntries | Where-Object {$_.ProjectID -eq "$Script:ClockifyProjectID"}
    $VictorTime | ForEach-Object {

        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime $_.TimeInterval.Duration
        $Script:VictorTotalTime = [decimal]$TimeSpent + [decimal]$VictorTotalTime

    }

    ####################
  
}

function GetProjectTaskLevelDetails {

    try{
        $Script:ClockifyProjectLevelDetails = Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/projects/" + $Script:ClockifyProjectID) -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Method Get
    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify project level details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Clockify project level details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
    }

    try{
        $ClockifyProjectTaskLevelDetails = Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/projects/" + $Script:ClockifyProjectID + "/tasks") -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Method Get
    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify task level details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Clockify task level details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
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

function LaunchTimeEntryModal {

    #The order of DDEs in this array needs to stay in this order. Changes this to this array would create a need to update the code (that sets the dropdown based on the existing DDE, if there is an existing DDE) directly above the ShowDialog for TimeEntryForm
    $DDELeads = @(
        "Angel",
        "Adam",
        "Howard",
        "Victor"
    )

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    $TimeEntryForm = New-Object System.Windows.Forms.Form
    $TimeEntryForm.Text = "Aktis Helper - Proposal Time Entry"
    $TimeEntryForm.Size = '600,400'
    $TimeEntryForm.StartPosition = "CenterScreen"
    $TimeEntryForm.MinimumSize = $TimeEntryForm.Size
    $TimeEntryForm.MaximizeBox = $False
    $TimeEntryForm.Topmost = $True

    $TimeEntryLabel = New-Object Windows.Forms.Label
    $TimeEntryLabel.Location = '5,5'
    $TimeEntryLabel.AutoSize = $True
    $TimeEntryLabel.Text = "Select the DDE lead and validate the following time data:"

    $ProposalNameLabel = New-Object Windows.Forms.Label
    $ProposalNameLabel.Location = '5,40'
    $ProposalNameLabel.AutoSize = $True
    $ProposalNameLabel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $ProposalNameLabel.Text = "$Script:ProposalName"

    $DDELeadDropDownBox = New-Object System.Windows.Forms.ComboBox
    $DDELeadDropDownBox.Location = New-Object System.Drawing.Size(195,100)
    $DDELeadDropDownBox.Size = New-Object System.Drawing.Size(65,20) 
    $DDELeadDDBoxLabel = New-Object Windows.Forms.Label
    $DDELeadDDBoxLabel.Location = New-Object System.Drawing.Point(95,102)
    $DDELeadDDBoxLabel.Text = "DDE Lead:"
    $DDELeadDDBoxLabel.ForeColor = 'Black'

    $TimeEntryLabelPC = New-Object Windows.Forms.Label
    $TimeEntryLabelPC.Location = '95,160'
    $TimeEntryLabelPC.AutoSize = $True
    $TimeEntryLabelPC.Text = "Proposal Coordination:"
    $TimeEntryListBoxPC = New-Object System.Windows.Forms.ListBox
    $TimeEntryListBoxPC.Location = New-Object System.Drawing.Point(220,160)
    $TimeEntryListBoxPC.Size = New-Object System.Drawing.Size(50,28)

    $TimeEntryLabelPW = New-Object Windows.Forms.Label
    $TimeEntryLabelPW.Location = '95,180'
    $TimeEntryLabelPW.AutoSize = $True
    $TimeEntryLabelPW.Text = "Proposal Writing:"
    $TimeEntryListBoxPW = New-Object System.Windows.Forms.ListBox
    $TimeEntryListBoxPW.Location = New-Object System.Drawing.Point(220,180)
    $TimeEntryListBoxPW.Size = New-Object System.Drawing.Size(50,28)

    $TimeEntryLabelDC = New-Object Windows.Forms.Label
    $TimeEntryLabelDC.Location = '95,200'
    $TimeEntryLabelDC.AutoSize = $True
    $TimeEntryLabelDC.Text = "Design Consulting:"
    $TimeEntryListBoxDC = New-Object System.Windows.Forms.ListBox
    $TimeEntryListBoxDC.Location = New-Object System.Drawing.Point(220,200)
    $TimeEntryListBoxDC.Size = New-Object System.Drawing.Size(50,28)

    $TimeEntryLabelIP = New-Object Windows.Forms.Label
    $TimeEntryLabelIP.Location = '95,220'
    $TimeEntryLabelIP.AutoSize = $True
    $TimeEntryLabelIP.Text = "Internal Project:"
    $TimeEntryListBoxIP = New-Object System.Windows.Forms.ListBox
    $TimeEntryListBoxIP.Location = New-Object System.Drawing.Point(220,220)
    $TimeEntryListBoxIP.Size = New-Object System.Drawing.Size(50,28)

    $TimeEntryLabelTotal = New-Object Windows.Forms.Label
    $TimeEntryLabelTotal.Location = '228,140'
    $TimeEntryLabelTotal.AutoSize = $True
    $TimeEntryLabelTotal.Text = "Total"

    $TimeEntryLabelAngel = New-Object Windows.Forms.Label
    $TimeEntryLabelAngel.Location = '280,140'
    $TimeEntryLabelAngel.AutoSize = $True
    $TimeEntryLabelAngel.Text = "Angel"
    $TimeEntryPCListBoxAngel = New-Object System.Windows.Forms.ListBox
    $TimeEntryPCListBoxAngel.Location = New-Object System.Drawing.Point(275,160)
    $TimeEntryPCListBoxAngel.Size = New-Object System.Drawing.Size(50,28)
    $TimeEntryPWListBoxAngel = New-Object System.Windows.Forms.ListBox
    $TimeEntryPWListBoxAngel.Location = New-Object System.Drawing.Point(275,180)
    $TimeEntryPWListBoxAngel.Size = New-Object System.Drawing.Size(50,28)
    $TimeEntryDCListBoxAngel = New-Object System.Windows.Forms.ListBox
    $TimeEntryDCListBoxAngel.Location = New-Object System.Drawing.Point(275,200)
    $TimeEntryDCListBoxAngel.Size = New-Object System.Drawing.Size(50,28)
    $TimeEntryIPListBoxAngel = New-Object System.Windows.Forms.ListBox
    $TimeEntryIPListBoxAngel.Location = New-Object System.Drawing.Point(275,220)
    $TimeEntryIPListBoxAngel.Size = New-Object System.Drawing.Size(50,28)

    $TimeEntryLabelAdam = New-Object Windows.Forms.Label
    $TimeEntryLabelAdam.Location = '337,140'
    $TimeEntryLabelAdam.AutoSize = $True
    $TimeEntryLabelAdam.Text = "Adam"
    $TimeEntryPCListBoxAdam = New-Object System.Windows.Forms.ListBox
    $TimeEntryPCListBoxAdam.Location = New-Object System.Drawing.Point(330,160)
    $TimeEntryPCListBoxAdam.Size = New-Object System.Drawing.Size(50,28)
    $TimeEntryPWListBoxAdam = New-Object System.Windows.Forms.ListBox
    $TimeEntryPWListBoxAdam.Location = New-Object System.Drawing.Point(330,180)
    $TimeEntryPWListBoxAdam.Size = New-Object System.Drawing.Size(50,28)
    $TimeEntryDCListBoxAdam = New-Object System.Windows.Forms.ListBox
    $TimeEntryDCListBoxAdam.Location = New-Object System.Drawing.Point(330,200)
    $TimeEntryDCListBoxAdam.Size = New-Object System.Drawing.Size(50,28)
    $TimeEntryIPListBoxAdam = New-Object System.Windows.Forms.ListBox
    $TimeEntryIPListBoxAdam.Location = New-Object System.Drawing.Point(330,220)
    $TimeEntryIPListBoxAdam.Size = New-Object System.Drawing.Size(50,28)

    $TimeEntryLabelVictor = New-Object Windows.Forms.Label
    $TimeEntryLabelVictor.Location = '391,140'
    $TimeEntryLabelVictor.AutoSize = $True
    $TimeEntryLabelVictor.Text = "Victor"
    $TimeEntryPCListBoxVictor = New-Object System.Windows.Forms.ListBox
    $TimeEntryPCListBoxVictor.Location = New-Object System.Drawing.Point(385,160)
    $TimeEntryPCListBoxVictor.Size = New-Object System.Drawing.Size(50,28)
    $TimeEntryPWListBoxVictor = New-Object System.Windows.Forms.ListBox
    $TimeEntryPWListBoxVictor.Location = New-Object System.Drawing.Point(385,180)
    $TimeEntryPWListBoxVictor.Size = New-Object System.Drawing.Size(50,28)
    $TimeEntryDCListBoxVictor = New-Object System.Windows.Forms.ListBox
    $TimeEntryDCListBoxVictor.Location = New-Object System.Drawing.Point(385,200)
    $TimeEntryDCListBoxVictor.Size = New-Object System.Drawing.Size(50,28)
    $TimeEntryIPListBoxVictor = New-Object System.Windows.Forms.ListBox
    $TimeEntryIPListBoxVictor.Location = New-Object System.Drawing.Point(385,220)
    $TimeEntryIPListBoxVictor.Size = New-Object System.Drawing.Size(50,28)

    $TimeEntryStatusBar = New-Object System.Windows.Forms.StatusBar
    $TimeEntryStatusBar.Text = "Ready to post..."

    $TimeEntrySubmitButton = New-Object System.Windows.Forms.Button
    $TimeEntrySubmitButton.Location = '221,275'
    $TimeEntrySubmitButton.Size = '100,23'
    $TimeEntrySubmitButton.Width = 125
    $TimeEntrySubmitButton.Text = "Post to Aktis"

    $TimeEntryRefreshTimeButton = New-Object System.Windows.Forms.Button
    $TimeEntryRefreshTimeButton.Location = '425,100'
    $TimeEntryRefreshTimeButton.Size = '100,23'
    $TimeEntryRefreshTimeButton.Width = 125
    $TimeEntryRefreshTimeButton.Text = "Refresh Time Data"

    $TimeEntryButton_Click = {

        $TimeEntryStatusBar.Text = "Posting data to Analytics Weekly Snapshot..."

        if($($DDELeadDropDownBox.SelectedItem.ToString()) -eq "Angel"){
            $Script:DDELead = "AP"
        }
        if($($DDELeadDropDownBox.SelectedItem.ToString()) -eq "Adam"){
            $Script:DDELead = "AG"
        }
        if($($DDELeadDropDownBox.SelectedItem.ToString()) -eq "Victor"){
            $Script:DDELead = "VA"
        }
        if($($DDELeadDropDownBox.SelectedItem.ToString()) -eq "Howard"){
            $Script:DDELead = "HB"
        }
        
        PostToAktis

    }

    $TimeRefreshButton_Click = {

        $TimeEntryStatusBar.Text = "Retrieving Clockify task level details..."
        GetProjectTaskLevelDetails

        $TimeEntryStatusBar.Text = "Retrieving Clockify user time data..."
        GetClockifyUserTimeEntryData

        $TimeEntryListBoxPW.Items.Clear()
        $TimeEntryListBoxPW.Items.Add([decimal]$Script:ProposalWritingTime)
        $TimeEntryListBoxPC.Items.Clear()
        $TimeEntryListBoxPC.Items.Add([decimal]$Script:ProposalCoordinationTime)
        $TimeEntryListBoxDC.Items.Clear()
        $TimeEntryListBoxDC.Items.Add([decimal]$Script:DesignConsultingTime)
        $TimeEntryListBoxIP.Items.Clear()
        $TimeEntryListBoxIP.Items.Add([decimal]$Script:InternalProjectTime)

        $TimeEntryPCListBoxAngel.Items.Clear()
        $TimeEntryPCListBoxAngel.Items.Add([decimal]$Script:AngelTotalPCTime)
        $TimeEntryPWListBoxAngel.Items.Clear()
        $TimeEntryPWListBoxAngel.Items.Add([decimal]$Script:AngelTotalPWTime)
        $TimeEntryDCListBoxAngel.Items.Clear()
        $TimeEntryDCListBoxAngel.Items.Add([decimal]$Script:AngelTotalDCTime)
        $TimeEntryIPListBoxAngel.Items.Clear()
        $TimeEntryIPListBoxAngel.Items.Add([decimal]$Script:AngelTotalIPTime)

        $TimeEntryPCListBoxAdam.Items.Clear()
        $TimeEntryPCListBoxAdam.Items.Add([decimal]$Script:AdamTotalPCTime)
        $TimeEntryPWListBoxAdam.Items.Clear()
        $TimeEntryPWListBoxAdam.Items.Add([decimal]$Script:AdamTotalPWTime)
        $TimeEntryDCListBoxAdam.Items.Clear()
        $TimeEntryDCListBoxAdam.Items.Add([decimal]$Script:AdamTotalDCTime)
        $TimeEntryIPListBoxAdam.Items.Clear()
        $TimeEntryIPListBoxAdam.Items.Add([decimal]$Script:AdamTotalIPTime)

        $TimeEntryPCListBoxVictor.Items.Clear()
        $TimeEntryPCListBoxVictor.Items.Add([decimal]$Script:VictorTotalPCTime)
        $TimeEntryPWListBoxVictor.Items.Clear()
        $TimeEntryPWListBoxVictor.Items.Add([decimal]$Script:VictorTotalPWTime)
        $TimeEntryDCListBoxVictor.Items.Clear()
        $TimeEntryDCListBoxVictor.Items.Add([decimal]$Script:VictorTotalDCTime)
        $TimeEntryIPListBoxVictor.Items.Clear()
        $TimeEntryIPListBoxVictor.Items.Add([decimal]$Script:VictorTotalIPTime)

        $TimeEntryStatusBar.Text = "Successfully refreshed Clockify user time data!"

    }

    $TimeEntrySubmitButton.Add_Click($TimeEntryButton_Click)
    $TimeEntryRefreshTimeButton.Add_Click($TimeRefreshButton_Click)

    $TimeEntryForm.SuspendLayout()
    $TimeEntryForm.Controls.Add($DDELeadDDBoxLabel)
    $TimeEntryForm.Controls.Add($DDELeadDropDownBox)
    $DDELeads | ForEach-Object{
        $DDELeadDropDownBox.Items.Add($_) | Out-Null
    }
    $TimeEntryForm.Controls.Add($TimeEntryLabel)
    $TimeEntryForm.Controls.Add($ProposalNameLabel)
    $TimeEntryForm.Controls.Add($TimeEntryLabelPW)
    $TimeEntryForm.Controls.Add($TimeEntryListBoxPW)
    $TimeEntryForm.Controls.Add($TimeEntryLabelPC)
    $TimeEntryForm.Controls.Add($TimeEntryListBoxPC)
    $TimeEntryForm.Controls.Add($TimeEntryLabelDC)
    $TimeEntryForm.Controls.Add($TimeEntryListBoxDC)
    $TimeEntryForm.Controls.Add($TimeEntryLabelIP)
    $TimeEntryForm.Controls.Add($TimeEntryListBoxIP)
    $TimeEntryForm.Controls.Add($TimeEntryStatusBar)
    $TimeEntryForm.Controls.Add($TimeEntrySubmitButton)
    $TimeEntryForm.Controls.Add($TimeEntryRefreshTimeButton)
    $TimeEntryForm.Controls.Add($TimeEntryLabelTotal)
    $TimeEntryForm.Controls.Add($TimeEntryLabelAngel)
    $TimeEntryForm.Controls.Add($TimeEntryPCListBoxAngel)
    $TimeEntryForm.Controls.Add($TimeEntryPWListBoxAngel)
    $TimeEntryForm.Controls.Add($TimeEntryDCListBoxAngel)
    $TimeEntryForm.Controls.Add($TimeEntryIPListBoxAngel)
    $TimeEntryForm.Controls.Add($TimeEntryLabelAdam)
    $TimeEntryForm.Controls.Add($TimeEntryPCListBoxAdam)
    $TimeEntryForm.Controls.Add($TimeEntryPWListBoxAdam)
    $TimeEntryForm.Controls.Add($TimeEntryDCListBoxAdam)
    $TimeEntryForm.Controls.Add($TimeEntryIPListBoxAdam)
    $TimeEntryForm.Controls.Add($TimeEntryLabelVictor)
    $TimeEntryForm.Controls.Add($TimeEntryPCListBoxVictor)
    $TimeEntryForm.Controls.Add($TimeEntryPWListBoxVictor)
    $TimeEntryForm.Controls.Add($TimeEntryDCListBoxVictor)
    $TimeEntryForm.Controls.Add($TimeEntryIPListBoxVictor)
    $TimeEntryForm.ResumeLayout()

    $TimeEntryListBoxPW.Items.Add([decimal]$Script:ProposalWritingTime)
    $TimeEntryListBoxPC.Items.Add([decimal]$Script:ProposalCoordinationTime)
    $TimeEntryListBoxDC.Items.Add([decimal]$Script:DesignConsultingTime)
    $TimeEntryListBoxIP.Items.Add([decimal]$Script:InternalProjectTime)

    $TimeEntryPCListBoxAngel.Items.Add([decimal]$Script:AngelTotalPCTime)
    $TimeEntryPWListBoxAngel.Items.Add([decimal]$Script:AngelTotalPWTime)
    $TimeEntryDCListBoxAngel.Items.Add([decimal]$Script:AngelTotalDCTime)
    $TimeEntryIPListBoxAngel.Items.Add([decimal]$Script:AngelTotalIPTime)

    $TimeEntryPCListBoxAdam.Items.Add([decimal]$Script:AdamTotalPCTime)
    $TimeEntryPWListBoxAdam.Items.Add([decimal]$Script:AdamTotalPWTime)
    $TimeEntryDCListBoxAdam.Items.Add([decimal]$Script:AdamTotalDCTime)
    $TimeEntryIPListBoxAdam.Items.Add([decimal]$Script:AdamTotalIPTime)

    $TimeEntryPCListBoxVictor.Items.Add([decimal]$Script:VictorTotalPCTime)
    $TimeEntryPWListBoxVictor.Items.Add([decimal]$Script:VictorTotalPWTime)
    $TimeEntryDCListBoxVictor.Items.Add([decimal]$Script:VictorTotalDCTime)
    $TimeEntryIPListBoxVictor.Items.Add([decimal]$Script:VictorTotalIPTime)

    if($Script:ExistingDDELead -eq "AP"){$DDELeadDropDownBox.SelectedItem = $DDELeadDropDownBox.Items[0]}
    if($Script:ExistingDDELead -eq "AG"){$DDELeadDropDownBox.SelectedItem = $DDELeadDropDownBox.Items[1]}
    if($Script:ExistingDDELead -eq "HB"){$DDELeadDropDownBox.SelectedItem = $DDELeadDropDownBox.Items[2]}
    if($Script:ExistingDDELead -eq "VA"){$DDELeadDropDownBox.SelectedItem = $DDELeadDropDownBox.Items[3]}
        
    $TimeEntryFormResult = $TimeEntryForm.ShowDialog()
        
    if($TimeEntryFormResult -eq [System.Windows.Forms.DialogResult]::OK){
        # Add OK action
    }
    elseif($TimeEntryFormResult -eq [System.Windows.Forms.DialogResult]::Cancel){
        $TimeEntryForm.Close()
        $TimeEntryForm.Dispose()
        $AktisPostConfForm.Close()
        $AktisPostConfForm.Dispose()
        $AktisHelperForm.Close()
        $AktisHelperForm.Dispose()
    }
    elseif(!($TimeEntryFormResult -eq [System.Windows.Forms.DialogResult]::OK)){
        $TimeEntryForm.Close()
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
    }    

}

function PostToAktis {

    if(!($Script:IsGratisRevision -eq $True)){

        if(($Script:ProposalVersion -eq "1") -or ($Script:ProposalVersion -eq "3") -or ($Script:ProposalVersion -eq "5") -or ($Script:ProposalVersion -eq "7") -or ($Script:ProposalVersion -eq "9") -or ($Script:ProposalVersion -eq "11") -or ($Script:ProposalVersion -eq "13")){
            $Script:IsGratisRevision = $False
        }
        elseif(($Script:ProposalVersion -eq "2") -or ($Script:ProposalVersion -eq "4") -or ($Script:ProposalVersion -eq "6") -or ($Script:ProposalVersion -eq "8") -or ($Script:ProposalVersion -eq "10") -or ($Script:ProposalVersion -eq "12") -or ($Script:ProposalVersion -eq "14")){
            $RevisionOrRewritePrompt = Show-Msgbox -Message "Is $Script:ProposalName a gratis revision?" -Icon "Information" -Button "YesNo" -Title "Aktis Helper"
            Switch ($RevisionOrRewritePrompt) {
            "Yes" {
                $Script:IsGratisRevision = $True
            }
            "No" {
                $Script:IsGratisRevision = $False
            }
            # "Cancel" {}
            }
        }
        
    }

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
        #Post to Proposal Statistics Sharepoint List
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Set-PnPListItem -List "Proposal Statistics" -Identity $Script:ProposalID -Values @{"ProposalVerison"="$Script:ProposalVersion"; "GratisRevision"="$Script:IsGratisRevision"; "ProposalExceedsSubscription_x002"="$Script:ProposalExceedsSubEnt"; "ProposalExceedsSubscription_x0020"="$Script:ProposalExceedsSubBasic"; "ProposedHours"="$Script:ProposedHours"; "HourlyRate"="$Script:HourlyRate"; "ProposedMRR"="$Script:TotalMRR"; "ProposedProfessionalServicesReve"="$Script:PSRevenue"; "DDELead"="$Script:DDELead"; "ProposedProductsRevenue"="$Script:ProductsRevenue"; "Proposed_x0020_Products_x0020_an"="$Script:TotalNRR"; "_x0031_00"="$Script:ProposalBiOpLT100"; "_x0031_000"="$Script:ProposalBiOpGT100"; "_x0032_00"="$Script:ProposalBiOpGT200"; "_x0033_00"="$Script:ProposalBiOpGT300"; "ProposalWritingTime"="$Script:ProposalWritingTime"; "ProposalCoordinationTime"="$Script:ProposalCoordinationTime"; "DesignConsultingTimeNew"="$Script:DesignConsultingTime"; "InternalProjectTime"="$Script:InternalProjectTime"}
    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to post Proposal Statistics data to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to post analytics data to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $TimeEntryForm.Close()
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
    }
    
    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Successfully posted Aktis Proposal Statistics and Client Analytics for $Script:ProposalName!`r`n`r`nProposal Number: $Script:ProposalNumber`r`nProposal Version: $Script:ProposalVersion`r`nProposed Hours: $Script:ProposedHours`r`nHourly Rate: $Script:HourlyRate`r`nProposed MRR: $Script:TotalMRR`r`nProposed Professional Services Revenue: $Script:PSRevenue`r`nProposed Products Revenue: $Script:ProductsRevenue`r`nTotal NRR: $Script:TotalNRR`r`nProposal Writing Time: $Script:ProposalWritingTime`r`nProposal Coordination Time: $Script:ProposalCoordinationTime`r`nDesign Consulting Time: $Script:DesignConsultingTime`r`nInternal Project Time: $Script:InternalProjectTime"
    $TimeEntryStatusBar.Text = "Successfully posted to Aktis..."
    $TimeEntryStatusBar.Text = "Posting data to Analytics Weekly Snapshot..."
    PostToLocalAnalyticsWeeklySnapshot

}

function PostToLocalAnalyticsWeeklySnapshot {

    function UpdateAWSDDENewProposal {

        <#
      
        if($($DDELeadDropDownBox.SelectedItem.ToString()) -eq "Angel"){
            $ExistingAPNewDrafts = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts).Value2
            $UpdatedAPNewDraftsValue = $ExistingAPNewDrafts + 1
            $DDEExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($DDEExistingNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                
            }
            else{
                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts) = "$UpdatedAPNewDraftsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts).AddComment("$DDEExistingNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts).Comment.Visible = $False
            }

            $ExistingTotals = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Value2
            $UpdatedTotalsValue = $ExistingTotals + 1
            $TotalsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($TotalsExistingNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                
            }
            else{
                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd) = "$UpdatedTotalsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).AddComment("$TotalsExistingNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Visible = $False
            }

            $ExcelSpreadsheetAWS.Save
            $ExcelSpreadsheetAWS.Close($True)
            $ExcelOpenAWS.Quit()
        }
        if($($DDELeadDropDownBox.SelectedItem.ToString()) -eq "Adam"){
            $ExistingAGNewDrafts = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts).Value2
            $UpdatedAGNewDraftsValue = $ExistingAGNewDrafts + 1
            $DDEExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($DDEExistingNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                
            }
            else{
                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts) = "$UpdatedAGNewDraftsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts).AddComment("$DDEExistingNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts).Comment.Visible = $False
            }

            $ExistingTotals = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Value2
            $UpdatedTotalsValue = $ExistingTotals + 1
            $TotalsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($TotalsExistingNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                
            }
            else{
                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd) = "$UpdatedTotalsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).AddComment("$TotalsExistingNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Visible = $False
            }

            $ExcelSpreadsheetAWS.Save
            $ExcelSpreadsheetAWS.Close($True)
            $ExcelOpenAWS.Quit()
        }
        if($($DDELeadDropDownBox.SelectedItem.ToString()) -eq "Victor"){
            $ExistingVANewDrafts = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts).Value2
            $UpdatedVANewDraftsValue = $ExistingVANewDrafts + 1
            $DDEExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($DDEExistingNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                
            }
            else{
                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts) = "$UpdatedVANewDraftsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts).AddComment("$DDEExistingNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts).Comment.Visible = $False
            }

            $ExistingTotals = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Value2
            $UpdatedTotalsValue = $ExistingTotals + 1
            $TotalsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($TotalsExistingNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                
            }
            else{
                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd) = "$UpdatedTotalsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).AddComment("$TotalsExistingNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Visible = $False
            }

            $ExcelSpreadsheetAWS.Save
            $ExcelSpreadsheetAWS.Close($True)
            $ExcelOpenAWS.Quit()
            
        }

        #>

        try{
            #Update DateProposalWritten column in Proposal Statistics Sharepoint List

            $TodaysDate = Get-Date

            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Set-PnPListItem -List "Proposal Queues" -Identity $Script:ProposalID -Values @{"DateProposalWritten"="$TodaysDate"}
        }
        catch{
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to post ProposalDateWritten to Proposal Statistics data to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
            [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
            [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to post ProposalDateWritten to Proposal Statistics data to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
            $TimeEntryForm.Close()
            $AktisPostConfForm.Close()
            $AktisHelperForm.Close()
        }

    }

    function UpdateAWSDDERevision {

        <#

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
            if($ExistingNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                
            }
            else{
                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd) = "$UpdatedGratisDraftsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd).AddComment("$ExistingNote`n$Script:ProposalNumber V$Script:ProposalVersion - $($DDELeadDropDownBox.SelectedItem.ToString())")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd).Comment.Visible = $False
            }

            $ExcelSpreadsheetAWS.Save
            $ExcelSpreadsheetAWS.Close($True)
            $ExcelOpenAWS.Quit()

        }

        #>

        try{
            #Update DateProposalWritten column in Proposal Statistics Sharepoint List

            $TodaysDate = Get-Date

            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Set-PnPListItem -List "Proposal Queues" -Identity $Script:ProposalID -Values @{"DateProposalWritten"="$TodaysDate"}
        }
        catch{
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to post ProposalDateWritten to Proposal Statistics data to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
            [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
            [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to post ProposalDateWritten to Proposal Statistics data to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
            $TimeEntryForm.Close()
            $AktisPostConfForm.Close()
            $AktisHelperForm.Close()
        }

    }
    
    function UpdateAWSPropCo {

        <#

        if($Script:ProposalVersion -eq '1'){

            if($Script:ProposedHours -lt "8"){
                #Proposal is a basic proposal
                $ExistingBasicCountColumn = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).Value2
                $UpdatedBasicCountValue = $ExistingBasicCountColumn + 1
                $ExistingTtlNewPrpslsDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Value2
                $UpdatedTtlNewPrpslsDelvrd = $ExistingTtlNewPrpslsDelvrd + 1
                $ExistingBasicCountNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                $ExistingTtlNewPrpslsDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                if(($ExistingBasicCountNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion") -or ($ExistingTtlNewPrpslsDelvrdNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion")){
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                    [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                    [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                    
                }
                else{
                    ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn)).Clear()
                    ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd)).Clear()
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn) = "$UpdatedBasicCountValue"
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).HorizontalAlignment = -4108
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).AddComment("$ExistingBasicCountNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd) = "$UpdatedTtlNewPrpslsDelvrd"
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).HorizontalAlignment = -4108
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).AddComment("$ExistingTtlNewPrpslsDelvrdNote`n$Script:ProposalNumber V$Script:ProposalVersion")
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
                if(($ExistingEnterpriseCountNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion") -or ($ExistingTtlNewPrpslsDelvrdNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion")){
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                    [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                    [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                    
                }
                else{
                    ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn)).Clear()
                    ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd)).Clear()
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn) = "$UpdatedEnterpriseCountValue"
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).HorizontalAlignment = -4108
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).AddComment("$ExistingEnterpriseCountNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd) = "$UpdatedTtlNewPrpslsDelvrd"
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).HorizontalAlignment = -4108
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).AddComment("$ExistingTtlNewPrpslsDelvrdNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                }
                $ExcelSpreadsheetAWS.Save
                $ExcelSpreadsheetAWS.Close($True)
                $ExcelOpenAWS.Quit()
            }
        }
        elseif($Script:ProposalVersion -gt '1'){

            if($Script:IsGratisRevision -eq $False){

                if($Script:ProposedHours -lt "8"){

                    $ExistingBillableRevsCmpltdDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).Value2
                    $UpdatedTtlBillableRevsCmpltdDelvrdCount = $ExistingBillableRevsCmpltdDelvrd + 1
                    $ExistingTtlBillableRevsCmpltdDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if($ExistingTtlBillableRevsCmpltdDelvrdNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd) = "$UpdatedTtlBillableRevsCmpltdDelvrdCount"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).AddComment("$ExistingTtlBillableRevsCmpltdDelvrdNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                    }

                    $ExistingBillableRevisionCount = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Value2
                    $UpdatedBillableRevisionCountValue = $ExistingBillableRevisionCount + 1
                    $BillableRevsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    
                    if($BillableRevsExistingNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn) = "$UpdatedBillableRevisionCountValue"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).AddComment("$BillableRevsExistingNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Comment.Visible = $False

                    }
    
                    $ExistingTtlNewPrpslsDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Value2
                    $UpdatedTtlNewPrpslsDelvrd = $ExistingTtlNewPrpslsDelvrd + 1
                    $ExistingTtlNewPrpslsDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if(($ExistingTtlNewPrpslsDelvrdNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion")){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd) = "$UpdatedTtlNewPrpslsDelvrd"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).AddComment("$ExistingTtlNewPrpslsDelvrdNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                    }

                    $ExistingBasicCountColumn = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).Value2
                    $UpdatedBasicCountValue = $ExistingBasicCountColumn + 1
                    $ExistingBasicCountNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if($ExistingBasicCountNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn) = "$UpdatedBasicCountValue"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).AddComment("$ExistingBasicCountNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                    }
    
                    $ExcelSpreadsheetAWS.Save
                    $ExcelSpreadsheetAWS.Close($True)
                    $ExcelOpenAWS.Quit()

                }
                if($Script:ProposedHours -ge "8"){

                    $ExistingBillableRevsCmpltdDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).Value2
                    $UpdatedTtlBillableRevsCmpltdDelvrdCount = $ExistingBillableRevsCmpltdDelvrd + 1
                    $ExistingTtlBillableRevsCmpltdDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if($ExistingTtlBillableRevsCmpltdDelvrdNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd) = "$UpdatedTtlBillableRevsCmpltdDelvrdCount"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).AddComment("$ExistingTtlBillableRevsCmpltdDelvrdNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                    }

                    $ExistingBillableRevisionCount = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Value2
                    $UpdatedBillableRevisionCountValue = $ExistingBillableRevisionCount + 1
                    $BillableRevsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    
                    if($BillableRevsExistingNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn) = "$UpdatedBillableRevisionCountValue"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).AddComment("$BillableRevsExistingNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Comment.Visible = $False

                    }
    
                    $ExistingTtlNewPrpslsDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Value2
                    $UpdatedTtlNewPrpslsDelvrd = $ExistingTtlNewPrpslsDelvrd + 1
                    $ExistingTtlNewPrpslsDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if(($ExistingTtlNewPrpslsDelvrdNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion")){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd) = "$UpdatedTtlNewPrpslsDelvrd"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).AddComment("$ExistingTtlNewPrpslsDelvrdNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                    }

                    $ExistingEnterpriseCountColumn = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).Value2
                    $UpdatedEnterpriseCountValue = $ExistingEnterpriseCountColumn + 1
                    $ExistingEnterpriseCountNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if($ExistingEnterpriseCountNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn) = "$UpdatedEnterpriseCountValue"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).AddComment("$ExistingEnterpriseCountNote`n$Script:ProposalNumber V$Script:ProposalVersion")
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
                    if($ExistingTtlGratisRevsCmpltdDelvrdNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlGratisRevsCmpltdDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlGratisRevsCmpltdDelvrd) = "$UpdatedTtlGratisRevsCmpltdDelvrdCount"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlGratisRevsCmpltdDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlGratisRevsCmpltdDelvrd).AddComment("$ExistingTtlGratisRevsCmpltdDelvrdNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                    }

                    $ExistingGratisRevisionCount = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn).Value2
                    $UpdatedGratisRevisionCountValue = $ExistingGratisRevisionCount + 1
                    $GratisRevsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    
                    if($GratisRevsExistingNote | Select-String "$Script:ProposalNumber V$Script:ProposalVersion"){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                        
                    }
                    else{
                    
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn) = "$UpdatedGratisRevisionCountValue"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn).AddComment("$GratisRevsExistingNote`n$Script:ProposalNumber V$Script:ProposalVersion")
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn).Comment.Visible = $False

                    }
    

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
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                #Write-Host "An error occurred:"
                #Write-Host $_.ScriptStackTrace
            }

            try{

                $CustomerProfileRawData = (Get-PnPListItem -List "Customer Profile" -Fields "Id","ClientName","EnterpriseSubscription","BasicSubscription","CurrentAgreementMRR","CurrentAgreementReproductionSubs","AdditionalMonthlyEnterpriseFee","AdditionalMonthlyBasicFee","ExpediteFee").FieldValues
            
                $Script:CustomerProfileData = @()
            
                $CustomerProfileRawData | ForEach-Object {
                    $Script:CustomerProfileData += $_
                }
        
            }
            catch{
                #Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Customer Profile data from SharePoint.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Customer Profile data from SharePoint.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
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
                Set-PnPListItem -List "Proposal Statistics" -Identity $Script:ProposalID -Values @{"ProposalExceedsSubscription_x002"="$Script:ProposalExceedsSubEnt"; "ProposalExceedsSubscription_x0020"="$Script:ProposalExceedsSubBasic"}
            }
            catch{
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to post ProposalExceedsSubscription to Proposal Statistics data to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to post ProposalExceedsSubscription to Proposal Statistics data to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
            }

            try{
                #Update DateProposalMarkedCompleted column in Proposal Statistics Sharepoint List
    
                $TodaysDate = Get-Date
    
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Set-PnPListItem -List "Proposal Statistics" -Identity $Script:ProposalID -Values @{"DateProposalMarkedCompleted"="$TodaysDate"}
            }
            catch{
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to post DateProposalMarkedCompleted to Proposal Statistics data to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to post DateProposalMarkedCompleted to Proposal Statistics data to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                $TimeEntryForm.Close()
                $AktisPostConfForm.Close()
                $AktisHelperForm.Close()
            }

        }

        #>
    }

    function ProposalCoordinatorCheck {
        $PropCoordinatorCheckPrompt = Show-Msgbox -Message "Did you just finish posting and delivering $Script:ProposalName and want to update the local Analytics Weekly Snapshot?" -Icon "Information" -Button "YesNo" -Title "Aktis Helper"
            Switch ($PropCoordinatorCheckPrompt) {
            "Yes" {UpdateAWSPropCo}
            "No" {}
            # "Cancel" {}
            }
    }

    <#

    $TimeEntryStatusBar.Text = "Posting to local Analytics Weekly Snapshot..."
    
    $LocalAnalyticsWeeklySnapshotFile = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Analytics\Analytics Weekly Snapshot.xlsx"

    try{

        if($(Test-FileLock $LocalAnalyticsWeeklySnapshotFile) -eq "True"){
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following analytics file is already open. Requesting to close file.`r`n$LocalAnalyticsWeeklySnapshotFile`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
            [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
            [void] [Microsoft.VisualBasic.Interaction]::MsgBox("$($LocalAnalyticsWeeklySnapshotFile | Split-Path -Leaf) is already open. Close the file and select OK.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
    
            while(
                $(Test-FileLock $LocalAnalyticsWeeklySnapshotFile) -eq "True"
            ){
                Start-Sleep -Milliseconds "500"
            }
        }

        $ExcelOpenAWS = New-Object -COMObject Excel.Application
        $ExcelOpenAWS.Visible = $True
        $ExcelSpreadsheetAWS = $ExcelOpenAWS.Workbooks.Open($LocalAnalyticsWeeklySnapshotFile)
        $ExcelSpreadsheetAWS.Worksheets(1).Activate()

    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to open Analytics Weekly Snapshot file. Ensure it's not already open.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to open Analytics Weekly Snapshot file. Ensure it's not already open", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $TimeEntryForm.Close()
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
        EXIT
        
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
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to determine current day.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to determine current day.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $TimeEntryForm.Close()
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
    }

    #>

    try{
        $DDEorPCPrompt = Show-Msgbox -Message "Are you the DDE that just finished writing $Script:ProposalName and want to update the local Analytics Weekly Snapshot?" -Icon "Information" -Button "YesNo" -Title "Aktis Helper"
        Switch ($DDEorPCPrompt) {
        "Yes" {
            if($Script:ProposalVersion -eq '1'){
                UpdateAWSDDENewProposal
            }
            elseif($Script:ProposalVersion -gt '1'){
                UpdateAWSDDERevision
            }
        }
        "No" {ProposalCoordinatorCheck}
        # "Cancel" {}
        }

    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to write to Analytics Weekly Snapshot.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to write to Analytics Weekly Snapshot.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        $TimeEntryForm.Close()
        $AktisPostConfForm.Close()
        $AktisHelperForm.Close()
    }

    $TimeEntryStatusBar.Text = "Successfully posted to local Analytics Weekly Snapshot!"
    $TimeEntryForm.Close()
    ThrowPostSuccessBalloon
    $AktisPostConfForm.Close()
    $AktisHelperForm.Close()

}

AuthenticateToSPO