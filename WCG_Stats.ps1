<# WCG_Stats.ps1

Chris Brown
24 May 20

This script requires a file named Daily_Results.txt
The first line must contain the date and time the statistics are for and be formatted like this:
Statistics Last Updated: 5/29/20 23:59:59 (UTC) [2 hour(s) ago]

The rest of the file will contain a record for each member's contribution for that day.

#>

Cls

$File    = "Daily_Results.txt"
$Header  = "Name,Joined,Years,Days,Hours,Minutes,Seconds,Points,Results,RecDate,Placement"
$InISE   = $Host.Name.Contains("ISE")
$Array   = @()
$Err     = $False
$Results = "Results"

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
# Program Start
#------------------------------------------------------------------------------------------

$ScriptPath = Get-ScriptDirectory
cd $ScriptPath
If ($Scriptpath[-1] -ne "\") {$ScriptPath = $ScriptPath+"\"}
$File = $ScriptPath+$File
$ResultFolder = $ScriptPath + $Results

If (-not (Test-Path $ResultFolder)) {
  New-Item -Path $ResultFolder -ItemType Directory -ErrorAction SilentlyContinue
  If (-not (Test-Path $ResultFolder)) {
    Write-Warning "Could Not Create $ResultFolder`n"
    If (-not $InISE) {Write-Host ""; Pause}
    Break
  }
}

If (-not (Test-Path $File)) {
  Write-Warning "Cannot Find File $File`n"
  If (-not $InISE) {Write-Host ""; Pause}
  Break
}

$Records = Get-Content $File
$D = $Records[0].Replace("Statistics Last Updated: ","")
$Date = $D -replace '\s.+$'
Try {
  $D = Get-Date ($Date) -Format "yyyy-MM-dd"
} Catch {
  $Err = $True
}

If ($Err) {
  Write-Warning "$File is not formatted correctly"
  Write-Host "The first line must contain a line like this: `"Statistics Last Updated: 5/23/20 23:59:59 (UTC) [18 hour(s) ago]`"`n"
  If (-not $InISE) {Write-Host ""; Pause}
  Break
}

$CSVFile = $ResultFolder+"\_$D.csv"

ForEach ($R in $Records) {
  $R = $R.Replace("`t`t",";")
  If (-not $R.Contains(" [")) {
    If ($R.Contains("- ")) {
      $R = $R.Replace("-"," -")
    }
    $Rec = $R -split ";"
    If ($Rec[1] -ne "") {
      $DT = $Rec[2] -split ":"
      $Item = [pscustomobject]@{
        Name      = $Rec[0]
        Joined    = $Rec[1] 
        Years     = $Dt[0]
        Days      = $Dt[1]
        Hours     = $Dt[2]
        Minutes   = $Dt[3]
        Seconds   = $Dt[4]   
        Points    = [long]$Rec[3] 
        Results   = [long]$Rec[4] 
        RecDate   = $D
        Placement = [int]0
      }
      $Array += $Item 
    }
  }
}

$Array = $Array | Sort Points -Descending

$I = 0
ForEach ($A in $Array) {
  $Array[$I].Placement = $I + 1
  $I++
}

$Array | Select-Object Name,Years,Days,Hours,Minutes,Seconds,Points,Results,Placement | ft -AutoSize

Write-Host "Saved to $CSVFile`n" -ForegroundColor Cyan
$Array | Export-Csv $CSVFile -NoTypeInformation

If (-not $InISE) {Write-Host ""; Pause}
