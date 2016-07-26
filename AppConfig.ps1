<#
                                                                    The Ark Church Fellowship One App config File

    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Future Info ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                http://stackoverflow.com/questions/24469265/running-a-powershell-script-restarting-and-then-continue-to-run
                powercfg -hibernate off for sleep function to work properly or else schedule task wont wake comp back up
#>

<#
################################ For future use: Recovering after a restart and continuing the script with error handling ###############################################

# This section must be the first section in the script. Set registry run once key to call script passing yes. Then well go into Error handling

Param ($Restart="") 

if ($Restart = "yes") {
    #[System.Windows.Forms.MessageBox]::Show("Passed: " + $Restart, "True")
} Else {
    #[System.Windows.Forms.MessageBox]::Show("Failed", "False")
}

#>

Add-Type -AssemblyName System.Windows.Forms # For Testing only

#
##################################################################### All Variables are Declared in this section. I try to use dynamic paths for portability.   #########
#

$Hour = (Get-Date).Hour.ToString()
#$Hour = "10"
$Day = (Get-Date).DayOfWeek
#$Day = "Sunday"
If ($Hour.Length -eq 2){}Else{$hour = "0" + $hour} # in AM we have a single digit so prepend Number 0
$Computer = $env:COMPUTERNAME.Substring(0,4) # Only use first 4 letters for computer matching in case statements
$LocalPath = $PSScriptRoot + "\"
$HourSS = $Hour + "SS.exe"
$SundaySpecial = "SunSpecial.exe"
$WednesdaySpecial = "WedSpecial.exe"
$LoftSpecial = "LoftSpecial.exe"
$HelpSS = "HelperSS.exe"

#
##################################################################### Set Screenaver variable based on the slides available in folder.   ################################
#


If (Test-Path ($LocalPath + $SundaySpecial)) {
    $SunSS = $SundaySpecial
} Else {
    $SunSS = $HourSS
}


If (Test-Path ($LocalPath + $WednesdaySpecial)) {
    $WedSS = $WednesdaySpecial
} Else {
    $WedSS = $HourSS
}


If (Test-Path ($LocalPath + $LoftSpecial)) {
    $LoftSS = $LoftSpecial
} Else {
    $LoftSS = $HourSS
}


#
#  ################################################################## Section for functions   ###########################################################################
#

Function Show-PopUp{
    [CmdletBinding()][OutputType([int])]Param( 
        [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("Msg")][string]$Message, 
        [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("Ttl")][string]$Title = $null, 
        [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("Duration")][int]$TimeOut = 0, 
        [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("But","BS")][ValidateSet( "OK", "OC", "AIR", "YNC" , "YN" , "RC")][string]$ButtonSet = "OK", 
        [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("ICO")][ValidateSet( "None", "Critical", "Question", "Exclamation" , "Information" )][string]$IconType = "None" 
         ) 
     
    $ButtonSets = "OK", "OC", "AIR", "YNC" , "YN" , "RC" 
    $IconTypes  = "None", "Critical", "Question", "Exclamation" , "Information" 
    $IconVals = 0,16,32,48,64 
    if((Get-Host).Version.Major -ge 3){ 
        $Button   = $ButtonSets.IndexOf($ButtonSet) 
        $Icon     = $IconVals[$IconTypes.IndexOf($IconType)] 
        } 
    else{ 
        $ButtonSets|ForEach-Object -Begin{$Button = 0;$idx=0} -Process{ if($_.Equals($ButtonSet)){$Button = $idx           };$idx++ } 
        $IconTypes |ForEach-Object -Begin{$Icon   = 0;$idx=0} -Process{ if($_.Equals($IconType) ){$Icon   = $IconVals[$idx]};$idx++ } 
        } 
    $objShell = New-Object -com "Wscript.Shell" 
    $objShell.Popup($Message,$TimeOut,$Title,$Button+$Icon) 
 
    <# 
        .SYNOPSIS 
            Creates a Timed Message Popup Dialog Box. 
 
        .DESCRIPTION 
            Creates a Timed Message Popup Dialog Box. 
 
        .OUTPUTS 
            The Value of the Button Selected or -1 if the Popup Times Out. 
            
            Values: 
                -1 Timeout   
                 1  OK 
                 2  Cancel 
                 3  Abort 
                 4  Retry 
                 5  Ignore 
                 6  Yes 
                 7  No 
 
        .PARAMETER Message 
            [string] The Message to display. 
 
        .PARAMETER Title 
            [string] The MessageBox Title. 
 
        .PARAMETER TimeOut 
            [int]   The Timeout Value of the MessageBox in seconds.  
                    When the Timeout is reached the MessageBox closes and returns a value of -1. 
                    The Default is 0 - No Timeout. 
 
        .PARAMETER ButtonSet 
            [string] The Buttons to be Displayed in the MessageBox.  
 
                     Values: 
                        Value     Buttons 
                        OK        OK                   - This is the Default           
                        OC        OK Cancel           
                        AIR       Abort Ignore Retry 
                        YNC       Yes No Cancel      
                        YN        Yes No              
                        RC        Retry Cancel        
 
        .PARAMETER IconType 
            [string] The Icon to be Displayed in the MessageBox.  
 
                     Values: 
                        None      - This is the Default 
                        Critical     
                        Question     
                        Exclamation  
                        Information  
             
        .EXAMPLE 
            $RetVal = Show-PopUp -Message "Data Trucking Company" -Title "Popup Test" -TimeOut 5 -ButtonSet YNC -Icon Exclamation 
 
        .NOTES 
            FunctionName : Show-PopUp 
            Created by   : Data Trucking Company 
            Date Coded   : 06/25/2012 16:55:46 
 
        .LINK 
             
     #> 
}

Function Close-App { # this function will Call the Show-Popup function in a loop, If no answer it will time out and shut down F1. No answer will wait 60 seconds

    $count = 0
    DO
    {
        $answer = Show-Popup -Message "***We are Shutting down Check-in***. If you are finished click YES. Otherwise Click NO and finish up" -Title "End Check-in" -TimeOut 6 -ButtonSet YN -IconType Information
        $count = $count + 1

        if ($answer -eq 6) #Yes
            {
             Stop-Process -Name "Fellow*"
             break;
            }
            elseif ($answer -eq 7)   #No
                 {
                   sleep -s 60
                 }
            elseif ($answer -eq -1) #Timed Out
                 {
                   sleep -s 5
                        if ($count -eq 2)
                            {
                            Stop-Process -Name "Fellow*"
                            Show-Popup -Message "***We are Shutting down Check-in***." -Title "End Check-in" -TimeOut 5 -ButtonSet OK -IconType Information
                            break;
                            }
                 }

    } While ($count -le 2)
}

Function Start-Setup {

runas administrator (get-process | ? { $_.mainwindowtitle -ne "" -and $_.processname -notlike "powershell*" } )| stop-process -Force

If (Get-Process | where-Object { $_.ProcessName -like "explorer"}) { & taskkill /im explorer.exe -f } Else {}

If (Get-Process | Where-Object { $_.ProcessName -like "OptionsWindow"}){
} Else { 
Start-Process OptionsWindow.exe -WorkingDirectory $LocalPath
}

$AppLoad = Start-Process AppLoad.exe -WorkingDirectory $LocalPath -PassThru

Start-Sleep -Seconds 1

    If (Get-Process | ? { $_.ProcessName -like $AppLoad.Name}){

        DO {
             Start-Sleep -Seconds 1
        } While ($AppLoad.HasExited -eq $false)

        If ($AppLoad.ExitCode -eq 1) {
               Start-Process explorer
               Stop-Process -Name OptionsWindow
               Start-Process $HelpSS -WorkingDirectory $LocalPath -WindowStyle Maximized
               Set-SSForeground ($HelpSS)
               Exit
        }
           
    } Else {
     If (Get-Process | where-Object { $_.ProcessName -like "explorer"}) { } Else {Start-Process explorer.exe}
    }
}

function Set-SSForeground([string]$SS){ # set the screensaver to front

Add-Type @"
  using System;
  using System.Runtime.InteropServices;
  public class Tricks {
     [DllImport("user32.dll")]
     [return: MarshalAs(UnmanagedType.Bool)]
     public static extern bool SetForegroundWindow(IntPtr hWnd);
  }
"@

    $h = (Get-Process $SS.Substring(0,$SS.Length-4)).MainWindowHandle
    [void] [Tricks]::SetForegroundWindow($h)
    sleep -sec 2
    $h2 = (Get-Process $SS.Substring(0,$SS.Length -4)).Id
    [void] [Tricks]::SetForegroundWindow($h2)
}

function Show-Process($Process, [Switch]$Maximize)
{
  $sig = '
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern int SetForegroundWindow(IntPtr hwnd);
  '
  
  if ($Maximize) { $Mode = 3 } else { $Mode = 4 }
  $type = Add-Type -MemberDefinition $sig -Name WindowAPI -PassThru
  $hwnd = $process.MainWindowHandle
  $null = $type::ShowWindowAsync($hwnd, $Mode)
  $null = $type::SetForegroundWindow($hwnd) 
}

#
# ################################################################### Main Setup Section ################################################################################
#
Switch ($Day)
{
        "Sunday"{ #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Sunday ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            If ($Hour -ile "08"){
                Switch ($Computer)
                    {
                        "Loft"{
                            Start-Setup
                        }
                        default{
                            Start-Setup
                            Start-Process $SunSS -WorkingDirectory $LocalPath
                            Set-SSForeground($SunSS)
                            Start-Sleep -Seconds (New-TimeSpan -End "8:00 Am").TotalSeconds # use this to get time until next service
                            Stop-Process -Name  $SunSS.Substring(0,$SunSS.Length-4)
                        }
                    }
            }
            If ($Hour -eq "09"){
                Switch ($Computer)
                    {
                        "Loft"{
                            Close-App
                            Start-Setup
                            Start-Process $LoftSS -WorkingDirectory $LocalPath
                            Set-SSForeground($LoftSS)
                            Start-Sleep -Seconds (New-TimeSpan -End "9:30 Am").TotalSeconds
                            Stop-Process -Name  $LoftSS.Substring(0,$LoftSS.Length-4)
                             }

                        default{
                            Close-App
                            Start-Setup
                            Start-Process $SunSS -WorkingDirectory $LocalPath
                            Set-SSForeground($SunSS)
                            Start-Sleep -Seconds (New-TimeSpan -End "9:30 Am").TotalSeconds
                            Stop-Process -Name  $SunSS.Substring(0,$SunSS.Length-4)
                            }
                    }
            }
            If ($Hour -eq "10"){
                 Switch ($Computer)
                    {
                        "Loft"{
                            Close-App
                            Start-Setup
                            Start-Process $LoftSS -WorkingDirectory $LocalPath
                            Set-SSForeground($LoftSS)
                            Start-Sleep -Seconds (New-TimeSpan -End "11:00 Am").TotalSeconds
                            Stop-Process -Name  $LoftSS.Substring(0,$LoftSS.Length-4)
                             }

                        default{
                            Close-App
                            Start-Setup
                            Start-Process $SunSS -WorkingDirectory $LocalPath
                            Set-SSForeground($SunSS)
                            Start-Sleep -Seconds (New-TimeSpan -End "11:00 Am").TotalSeconds
                            Stop-Process -Name  $SunSS.Substring(0,$SunSS.Length-4)
                            }
                    }
            
            }
            If ($Hour -eq "12"){
                 Switch ($Computer)
                    {
                        "Loft"{
                            Close-App
                            Start-Process $LoftSS -WorkingDirectory $LocalPath
                            Start-Sleep -Seconds (New-TimeSpan -End "12:45 Pm").TotalSeconds
                            Stop-Process -Name  $LoftSS.Substring(0,$SunSS.Length-4)
                            & nircmd.exe monitor off
                            #& cmd.exe /c %windir%\System32\rundll32.exe powrprof.dll,SetSuspendState Standby    
                        }
                        default{
                            Close-App
                            Start-Process $SunSS -WorkingDirectory $LocalPath
                            Start-Sleep -Seconds (New-TimeSpan -End "12:45 Pm").TotalSeconds
                            Stop-Process -Name  $SunSS.Substring(0,$SunSS.Length-4)
                            & nircmd.exe monitor off
                            #& cmd.exe /c %windir%\System32\rundll32.exe powrprof.dll,SetSuspendState Standby
                            }
                    }
            }
        } # Sunday case end

        "Tuesday"{ #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Tuesday ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            If ($Hour -ile "16"){
                Switch ($Computer)
                    {
                        default{Start-Setup }
                    }
            }

            If ($Hour -eq "17"){
                 Switch ($Computer)
                    {
                        default{Start-Setup}
                    }
            }
        } # Tuesday case end

        "Wednesday"{ #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Wednesday ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            if ($Hour -ile "16"){
                Switch ($Computer){
                    default{ Start-Setup}
            
            }
            }
            
            
            If ($Hour -eq "17"){
                 Switch ($Computer)
                    {
                        "Loft"{
                            Start-Setup
                             }
						"Sout"{
							Start-Setup
						}
                        default{
                            Start-Process $WedSS -WorkingDirectory $LocalPath
                            }
                    }
            
            }
            If ($Hour -eq "18"){
                 Switch ($Computer)
                    {
                        "Loft"{} # we want to catch this and do nothing
						
						"Sout"{}

                        default{
                            Close-App
                            Start-Setup
                            $scr = Start-Process $WedSS -WorkingDirectory $LocalPath -PassThru
                            Start-Sleep -Seconds 1
                            Set-SSForeground($WedSS)
                            Show-Process -Process $scr -Maximize
                            Start-Sleep -Seconds (New-TimeSpan -End "6:30 Pm").TotalSeconds
                            Stop-Process -Name  $SunSS.Substring(0,$SunSS.Length-4)
                            }
                    }
            }
            If ($Hour -eq "19"){
                Switch ($Computer)
                    {
                        "Loft"{
                            Close-App
                            Start-Process $LoftSS -WorkingDirectory $LocalPath
                            Set-SSForeground($LoftSS)
                            Start-Sleep -Seconds (New-TimeSpan -End "8:15 Pm").TotalSeconds
                            Stop-Process $LoftSS.Substring(0,$LoftSS.Length-4)
                            & nircmd.exe monitor off
                            #& %windir%\System32\rundll32.exe powrprof.dll,SetSuspendState Standby
                            }
                        default{
                            Close-App
                            Start-Process $WedSS -WorkingDirectory $LocalPath
                            Set-SSForeground($WedSS)
                            Start-Sleep -Seconds (New-TimeSpan -End "8:15 Pm").TotalSeconds
                            Stop-Process $WedSS.Substring(0,$WedSS.Length-4)
                            & nircmd.exe monitor off
                            #& %windir%\System32\rundll32.exe powrprof.dll,SetSuspendState Standby
                            }

                    }
        
            }
        } # Wednesday case end

        "Friday" { #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Friday ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            If ($Hour -eq "18"){
                 Switch ($Computer)
                    {
                        default{Start-Setup}
                    }
            }
        } # friday case end
}