<# Results.ps1

Chris Brown
24 May 20

See the readme.txt file for a description of this script.

#>

Cls

Remove-Variable * -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Windows.Forms

$Results  = "Results"
$Header   = "Name,Joined,Years,Days,Hours,Minutes,Seconds,Points,Results,RecDate,Placement,Change"
$InISE    = $Host.Name.Contains("ISE")
$CSVFile  = "Results.csv"
$TxtFile  = "Results.txt"
$HTMFile  = "Results.log"
$Monthly  = $True # Updates the header to display Monthly Report
$ShowNP   = $True # Show accounts with no progress
$Array    = @()

#------------------------------------------------------------------------------------------
# This function gets the current directory that the script is saved in
#------------------------------------------------------------------------------------------
Function Get-ScriptDirectory
{
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value;
    if($Invocation.PSScriptRoot) {
      $Invocation.PSScriptRoot;
    } Elseif($Invocation.MyCommand.Path) {
      Split-Path $Invocation.MyCommand.Path
    } else {
      $Invocation.InvocationName.Substring(0,$Invocation.InvocationName.LastIndexOf("\"));
    }
}

#------------------------------------------------------------------------------------------
# This function Writes Information to the html file
#------------------------------------------------------------------------------------------
Function Write-HTM
{
  Param ($String) 
  $String | Out-File $HTMFile -Encoding utf8 -Append
}

#------------------------------------------------------------------------------------------
# Program Start
#------------------------------------------------------------------------------------------

$ScriptPath = Get-ScriptDirectory
cd $ScriptPath
If ($Scriptpath[-1] -ne "\") {$ScriptPath = $ScriptPath+"\"}
$ResultFolder = $ScriptPath + $Results


$Ans = ""
Do {
  Write-Host "Is this a Monthly Report? " -NoNewline -ForegroundColor Cyan
  Write-Host "(" -NoNewline
  Write-Host "Y" -NoNewline -ForegroundColor Cyan
  Write-Host "/" -NoNewline
  Write-Host "N" -NoNewline -ForegroundColor Cyan
  Write-Host "/" -NoNewline
  Write-Host "Q" -NoNewline -ForegroundColor Cyan
  Write-Host ") " -NoNewline
  $Ans = (Read-Host).ToUpper()
} Until ($Ans -eq "Y" -or $Ans -eq "N" -or $Ans -eq "Q")
If ($Ans -eq "Q") {
  Write-Host "Exiting`n"
  If (-not $InISE) {Write-Host ""; Pause}
  Break
}
$Monthly = ($Ans -eq "Y")


If (-not $Monthly) {
  $Ans = ""
  Do {
    Write-Host "Display People With No Progress? " -NoNewline -ForegroundColor Cyan
    Write-Host "(" -NoNewline
    Write-Host "Y" -NoNewline -ForegroundColor Cyan
    Write-Host "/" -NoNewline
    Write-Host "N" -NoNewline -ForegroundColor Cyan
    Write-Host "/" -NoNewline
    Write-Host "Q" -NoNewline -ForegroundColor Cyan
    Write-Host ") " -NoNewline
    $Ans = (Read-Host).ToUpper()
  } Until ($Ans -eq "Y" -or $Ans -eq "N" -or $Ans -eq "Q")
  If ($Ans -eq "Q") {
    Write-Host "Exiting`n"
    If (-not $InISE) {Write-Host ""; Pause}
    Break
  }
  $ShowNP = ($Ans -eq "Y")
}

# Open the first file
If ($Monthly) {
  $Title = "Select the Last File of the Month: "
} Else {
  $Title = "Select the Newest File: "
}

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
  InitialDirectory = $ResultFolder
  Multiselect      = $False
  DefaultExt       = '*.csv'
  Filter           = 'CSV Files (*.csv)|_*.csv|All Files (*.*)|*.*'
  Title            = $Title
}

Write-Host $Title -ForegroundColor Cyan -NoNewline

If ($FileBrowser.ShowDialog() -eq "Cancel") {
  Write-Host "Exiting`n"
  $FileBrowser.Dispose()
  Break
}
$File1  = $FileBrowser.FileName
$Folder = $File1 | Split-Path
Write-Host $File1
$FileBrowser.Dispose()

# Open the Second file
If ($Monthly) {
  $Title = "Select the First File of the Month: "
} Else {
  $Title = "Select the Previous File: "
}

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
  InitialDirectory = $Folder
  Multiselect      = $False
  DefaultExt       = '*.csv'
  Filter           = 'CSV Files (*.csv)|_*.csv|All Files (*.*)|*.*'
  Title            = $Title
}

Write-Host $Title -ForegroundColor Cyan -NoNewline
If ($FileBrowser.ShowDialog() -eq "Cancel") {
  Write-Host "Exiting`n"
  $FileBrowser.Dispose()
  Break
}
$File2 = $FileBrowser.FileName
Write-Host $File2
$FileBrowser.Dispose()

$InFile1 = Import-Csv $File1
$InFile2 = Import-Csv $File2
Write-Host

$StatDate    = $InFile1[0].RecDate
$MonthlyDate = $StatDate.Substring(0,7)
$CSVFile     = $ResultFolder + "\$StatDate`_$CSVFile"
$TxtFile     = $ResultFolder + "\$StatDate`_$TxtFile"
$HTMFile     = $ResultFolder + "\$StatDate`_$HTMFile"

ForEach ($In1 in $InFile1) {
  If ($In1.Retired -eq "No") {
  
  $Change    = 0
  $PointsInc = 0
  $ResultInc = 0
  $Found = $False
  $Tabs = "`t"
  If ($In1.Name.Length -le 17) {$Tabs = "`t`t"}
  If ($In1.Name.Length -le 15) {$Tabs = "`t`t`t"}
  If ($In1.Name.Length -le 11) {$Tabs = "`t`t`t`t"}
  If ($In1.Name.Length -le 7)  {$Tabs = "`t`t`t`t`t"}
  If ($In1.Name.Length -le 3)  {$Tabs = "`t`t`t`t`t`t"}
  Write-Host $In1.Name -ForegroundColor Cyan -NoNewline
  Write-Host "$Tabs= " -NoNewline
  Write-Host $In1.Placement -ForegroundColor Cyan -NoNewline
  $Space = "  "
  If ($In1.Placement.Length -eq 1) {$Space = "   "}
  Write-Host $Space -NoNewline
  
  ForEach ($In2 in $InFile2) {
    If ($In2.Name -eq $In1.Name) {
      $Found = $True
      $FG = "Cyan"

      If ($In2.Placement -lt $In1.Placement) {
        $FG = "Yellow"
        $Change = [int]$In2.Placement - [int]$In1.Placement
      }

      If ($In2.Placement -gt $In1.Placement) {
        $FG = "Green"
        $Change = [int]$In2.Placement - [int]$In1.Placement
      }

      Write-Host $In2.Placement -ForegroundColor $FG -NoNewline
      $Space = " "
      If ($In2.Placement.Length -eq 1) {$Space = "  "}
      Write-Host $Space -NoNewline

      If ($Change -eq 0) { $Num = $Null } Else { $Num = $Change }
      If ([int]$Change -gt 0) { $Num = "+" + $Num }
      Write-Host "  $Num" -ForegroundColor $FG -NoNewline

      $PointsInc = '{0:N0}' -f ([long]$In1.Points - [long]$In2.Points)
      Write-Host "  $PointsInc" -NoNewline

      $ResultInc = '{0:N0}' -f ([long]$In1.Results - [long]$In2.Results)
      Write-Host "  $ResultInc"
      Break
    }
  }
  If (-not $Found) {
    $Num = "New"
    Write-Host "New`t" -ForegroundColor Magenta -NoNewline
    $PointsInc = '{0:N0}' -f ([long]$In1.Points)
    Write-Host " $PointsInc" -NoNewline

    $ResultInc = '{0:N0}' -f ([long]$In1.Results)
    Write-Host "  $ResultInc"
  }

  If ($ShowNP) {
    $SaveRec = $True
  } Else {
    $SaveRec = [long]$PointsInc -gt 0
  }

  If ($SaveRec) {
    $Item = [pscustomobject]@{
      Name       = $In1.Name
      Joined     = $In1.Joined 
      Years      = $In1.Years
      Days       = $In1.Days
      Hours      = $In1.Hours
      Minutes    = $In1.Minutes
      Seconds    = $In1.Seconds
      Points     = $In1.Points
      Results    = $In1.Results
      RecDate    = $In1.RecDate
      Rank       = $In1.Placement
      PointsInc  = $PointsInc 
      ResultsInc = $ResultInc
      POSChange  = $Num
    }
    $Array += $Item 
  }
}

}

$Array | Export-Csv $CSVFile -NoTypeInformation

$I = 0
$Array | ForEach {
  If ($_.Name[0] -eq " ") {
    $Array[$I].Name    = $_.Name.Trim()
  }
  $Array[$I].Points  = '{0:N0}' -f [long]$_.Points 
  $Array[$I].Results = '{0:N0}' -f [long]$_.Results 
  $I++
}

Write-Host "`nCreating File: " -NoNewline
Write-Host $HTMFile -ForegroundColor Cyan
If ($Monthly) {
  "[font=courier new][code][b][size=3]Monthly Results for $MonthlyDate[/size] (Retired Accounts Not Displayed)[/b]" | Out-File $HTMFile -Encoding utf8
} Else {
  "[font=courier new][code][b][size=3]Results for $StatDate[/size] (Retired Accounts Not Displayed)[/b]" | Out-File $HTMFile -Encoding utf8
}

Write-HTM "Rank`tName`t`t`tPoints`t`tPoints Inc`tResults Inc`tRank Change"
Write-HTM "-------`t-----------------------`t---------------`t---------------`t---------------`t-----------"

ForEach ($A in $Array) {

  If ($A.Name.Length -le 23) {$NameTabs = "`t"}
  If ($A.Name.Length -le 15) {$NameTabs = "`t`t"}
  If ($A.Name.Length -le 7)  {$NameTabs = "`t`t`t"}
  
  If ([long]$A.Points    -ge 1000000) { $PointTabs = "`t" } Else {$PointTabs = "`t`t"}
  If ([long]$A.Pointsinc -ge 1000000) { $IncTabs   = "`t" } Else {$IncTabs   = "`t`t"}

  Write-HTM "$($A.Rank)`t$($A.Name)$NameTabs$($A.Points)$PointTabs$($A.PointsInc)$IncTabs$($A.ResultsInc)`t`t$($A.PosChange)"
}

If (-not $ShowNP) {
  Write-HTM " "
  Write-HTM "(Showing Only Active Members)"
}
Write-HTM "[/code][/font]"
Notepad.exe $HTMFile

If (-not $InISE) {Write-Host ""; Pause}
