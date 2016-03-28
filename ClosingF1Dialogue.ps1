<# 
    10/27/15
    this is for stopping F1 at the end of a service, and starting the screen saver based on user input.
#>


[CmdletBinding()][OutputType([int])]Param( 
    [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("Msg")][string]$Message, 
    [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("Ttl")][string]$Title = $null, 
    [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("Duration")][int]$TimeOut = 0, 
    [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("But","BS")][ValidateSet( "OK", "OC", "AIR", "YNC" , "YN" , "RC")][string]$ButtonSet = "OK", 
    [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("ICO")][ValidateSet( "None", "Critical", "Question", "Exclamation" , "Information" )][string]$IconType = "None" 
     ) 
 
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
    $objShell.Popup($Message,$TimeOut,$Title,$Button+$Icon+vbSystemModal) 
 
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
$count = 0

DO
{
    $answer = Show-Popup -Message "***We are Shutting down Check-in***. If you are finished click YES. Otherwise Click NO and finish up" -Title "End Check-in" -TimeOut 10 -ButtonSet YN -IconType Information
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
                    if ($count -eq 4)
                        {
                        Stop-Process -Name "Fellow*"
                        Show-Popup -Message "***We are Shutting down Check-in***." -Title "End Check-in" -TimeOut 10 -ButtonSet OK -IconType Information
                        break;
                        }
             }

} While ($count -le 3)
