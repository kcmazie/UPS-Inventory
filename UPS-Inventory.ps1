Param(
    [switch]$Console = $false,                                                  #--[ Set to true to enable local console result display. Defaults to false ]--
    [switch]$Debug = $False,                                                    #--[ Generates extra console output for debugging.  Defaults to false ]--
    [switch]$EnableExcel = $True                                                #--[ Defaults to use Excel. ]--  
)
<#==============================================================================
         File Name : UPS-Inventory.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 
       Description : Uses SNMP to poll and track APC UPS devices using MS Excel.
                   : 
             Notes : Normal operation is with no command line options.  Commandline options noted below.
                   :
      Requirements : Requires the PowerShell SNMP library from https://www.powershellgallery.com/packages/SNMPv3
                   : Currently designed to poll APC UPS devices.  UPS NMC must have SNMP v3 active.
                   : Script checks for active SNMPv1, FTP, and SNMPv3.
                   : Will generate a new spreadsheet if none exists by using a text file located in the same folder
                   : as the script, one IP per line.  Default operation is to check for text file first, then if not
                   : found check for an existing spreadsheet also in the same folder.  If an existing spreadsheet
                   : is located the target list is compliled from column A.  It will also copy a master spreadsheet
                   : to a working copy that gets processed.  Up to 10 backup copies are retained prior to writing
                   : changes to the working copy.
                   : 
          Warnings : Excel is set to be visible (can be changed) so don't mess with it while the script is running or it can crash.
                   :   
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   :
           Credits : Code snippets and/or ideas came from many sources including but 
                   :   not limited to the following:
                   : 
    Last Update by : Kenneth C. Mazie                                           
   Version History : v1.0 - 08-16-22 - Original 
    Change History : v2.0 - 09-00-22 - Numerous operational & bug fixes prior to v3.
                   : v3.1 - 09-15-22 - Cleaned up final version for posting.
                   : v4.0 - 04-12-23 - Too many changes to list
                   : v4.1 - 07-03-23 - Added age and LDOS dates. 
                   : v5.0 - 01-17-24 - Fixed DNS lookup.  Fixed last test result.  Fixed color coding of hostname for
                   :                   numerous events.  Added hostname cell comments to describe color coding.
                   :                  
==============================================================================#>

#Requires -version 5
##equires -Modules @{ ModuleName="SNMPv3"; ModuleVersion="1.1.1" }
Clear-Host 

#--[ Variables ]---------------------------------------------------------------
$DateTime = Get-Date -Format MM-dd-yyyy_HHmmss 
$Today = Get-Date -Format MM-dd-yyyy 
$IPTextFile = "IPList.txt"
$Script:v3UserTest = $False
$SafeUpdate = $True
$ExcelWorkingCopy = ($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xlsx"
$ConfigFile = ($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xml"
$TestFileName = "$PSScriptRoot\TestFile.txt"

#--[ The following can be hardcoded here or loaded from the XML file ]--
#$SourcePath = < See external config file >
#$ExcelSourceFile = < See external config file >
#$SMNPv3User = < See external config file >
#$SMNPv3AltUser = < See external config file >
#$SNMPv3Secret = < See external config file >

#--[ Runtime tweaks for testing ]--
$EnableExcel = $True
$Console = $True
$Debug = $false #true
$CloseOpen = $true

#------------------------------------------------------------
#
$erroractionpreference = "stop"
try{
    if (!(Get-Module -Name SNMPv3)) {
        Get-Module -ListAvailable SNMPv3 | Import-Module | Out-Null
    }
    Install-Module -Name SNMPv3
}Catch{

}

#==[ Functions ]===============================================================
Function StatusMsg ($Msg, $Color){
    If ($Null -eq $Color){
        $Color = "Magenta"
    }
    Write-Host "-- Script Status: $Msg" -ForegroundColor $Color
    $Msg = ""
}

function RGB ($red, $green, $blue ){
    return [System.Int]($red + $green * 256 + $blue * 256 * 256)
}

Function OctetString2String ($Result){
    $Bytes = [System.Text.Encoding]::Unicode.GetBytes($Result)
    $SaveVal = "" 
    ForEach ($Value in $Bytes){
        If ($Value -ne " "){
            $SaveVal += ([System.Text.Encoding]::ASCII.GetString($Value)).trim()                
        }
    }  
    Return $SaveVal
}

Function LoadConfig {
    #--[ Read and load configuration file ]-------------------------------------
    if (!(Test-Path "$PSScriptRoot\$ConfigFile")){                       #--[ Error out if configuration file doesn't exist ]--
        StatusMsg "MISSING CONFIG FILE.  Script aborted." " Red"
        break;break;break
    }else{
        [xml]$Configuration = Get-Content "$PSScriptRoot\$ConfigFile"  #--[ Read & Load XML ]--    
        $Script:SourcePath = $Configuration.Settings.General.SourcePath
        $Script:ExcelSourceFile = $Configuration.Settings.General.ExcelSourceFile
        $Script:DNS = $Configuration.Settings.General.DNS 
        $Script:SMNPv3User = $Configuration.Settings.Credentials.SMNPv3User
        $Script:SMNPv3AltUser = $Configuration.Settings.Credentials.SMNPv3AltUser
        $Script:SNMPv3Secret = $Configuration.Settings.Credentials.SMNPv3Secret
        $Script:PasswordFile = $Configuration.Settings.Credentials.PasswordFile
        $Script:KeyFile = $Configuration.Settings.Credentials.KeyFile
    }
}

Function Write2Excel ($WorkSheet,$Row,$Col,$Data,$Format,$Debug){
    $Existing = $WorkSheet.Cells.Item($Row,$Col).Text                               #--[ Read existing spreadsheet cell data for comparison ]-- 
#    $Index = $WorkSheet.Cells($Row-2,1).Interior.ColorIndex                         #--[ Determine existing row cell color index ]--
#    $Anomaly = ""                                                                   #--[ A return flag for doing extra adjustments ]--
   # sleep -sec .5
   # write-host $Col -ForegroundColor red
   # $debug = $true

    If ($Debug){
        write-host "`n  New Data      :"$Data -ForegroundColor Green
        write-host "  Existing Data :"$Existing -ForegroundColor Cyan
    }
    If ($Script:SpreadSheet -eq "New"){                                             #--[ Creating a new spreadsheet, set all cells to black ]--
        $Worksheet.Cells($Row, $Col).Font.Bold = $False
        $Worksheet.Cells($Row, $Col).Font.ColorIndex = 0                            #--[ Black ]--    
    }Else{                                                                          #--[ Using existing spreadsheet. ]--      
        If ($Existing -eq ""){
            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 0                        #--[ Black to denote a new item ]--
        }Else{
            Switch ($Format){
                "existing" {
                    $Data = $Existing                                               #--[ Never over-write the existing data ]--
                }
                "number" {
                    If ($Existing -gt $Data){
                        $Worksheet.Cells($Row, $Col).Font.Bold = $false
                        $Worksheet.Cells($Row, $Col).Font.ColorIndex = 10           #--[ Green ]--
                    }
                    If ($Data -gt $Existing){
                        $Worksheet.Cells($Row, $Col).Font.Bold = $true
                        $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3            #--[ Red ]--
                    }
                }
                "date" {
                    $Worksheet.Cells($Row, $Col).NumberFormat = "mm/dd/yyyy"
                    $Worksheet.Cells($Row, $Col).Font.Bold = $False
                    $Worksheet.Cells($Row, $Col).Font.ColorIndex = 0                #--[ Black ]--
                }
                "run" {                                                             #--[ Calculate Batt Runtime and load deltas ]--
                    If (($Data -like "*NoSuch*") -or ($WorkSheet.Cells.Item($Row,$Col).Text -ne "UPS")){
                        # --[ Do Nothing ]--
                    }Else{
                        If ($Col -eq 31){  
                            #--[ NOTE !!! ]--                                         
                            #--[ Column 31 is "AE" which should equate to Battery Load % ]--                        
                            #--[ Adjust this if column count is altered ]--
                            $Existing = $Existing.Split(" ")[0]
                            $Delta = $Existing-$Data
                            If ([int]$Data -gt [int]$Existing){                     #--[ Load increase ]--
                                $Worksheet.Cells($Row, $Col).Font.Bold = $True
                                $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3    #--[ Red ]--
                                $Data = $Data+"  (+"+(-$Delta)+")"
                            }ElseIf ([int]$Existing -gt [int]$Data){                #--[ Load decrease ]-- 
                                $Worksheet.Cells($Row, $Col).Font.Bold = $True
                                $Worksheet.Cells($Row, $Col).Font.ColorIndex = 10   #--[ Green ]--
                                $Data = $Data+"  (-"+($Delta)+")"
                            }
                        }Else{                                                      
                            #--[ If column is NOT 31, then col 28 "AB" gets processed ]--
                            $DeltaH = ([int]$Existing.Split(" ")[0] - [int]$Data.Split(" ")[0])*60   
                            $DeltaM = [int]$Existing.Split(" ")[2] - [int]$Data.Split(" ")[2]
                            $Delta = $DeltaH+$DeltaM                   
                            If ($Delta -lt 0){                                      #--[ Runtime gain ]--
                                $Worksheet.Cells($Row, $Col).Font.Bold = $True
                                $Worksheet.Cells($Row, $Col).Font.ColorIndex = 10   #--[ Green ]--
                                $Data = $Data+"  (+"+(-$Delta)+")"
                            }ElseIf($Delta -gt 0){                                  #--[ Runtime loss ]--
                                $Worksheet.Cells($Row, $Col).Font.Bold = $True
                                $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3    #--[ Red ]--
                                $Data = $Data+"  (-"+$Delta+")"
                            }
                        }                        
                        If ($Delta -eq 0){
                            $Worksheet.Cells($Row, $Col).Font.Bold = $False
                            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 0        #--[ Black ]--
                        }
                    }
                }
                "mac" {                                                             #--[ Process MAC address(es) ]--
                    If (($Existing -eq $Obj.RealMAC) -Or ($Existing -eq $Data)){    #--[ Exisitng matches all data ]--
                        $Worksheet.Cells($Row, $Col).Font.Bold = $true
                        $Worksheet.Cells($Row, $Col).Font.ColorIndex = 10           #--[ Flag as bold green. ]--
                        If ($Null -ne $Worksheet.Cells($Row, $Col).Comment()){  
                            $Worksheet.Cells($Row, $Col).Comment.Delete()           #--[ Remove any cell comment ]--
                        }
                        $Data = $Existing
                    }ElseIf (($Data -like "*:3f:*") -Or ($Existing -like "*:3f:*")){  #--[ 3F in position 3 and/or 6 denotes bad data ]--
                        If (($Obj.RealMAC -eq "Not Detected") -Or ($Obj.RealMAC -like "*:FF:")){
                            $Worksheet.Cells($Row, $Col).Font.Bold = $False
                            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3        #--[ Red ]--
                            If ($Null -ne $Worksheet.Cells($Row, $Col).Comment()){  #--[ If a previous comment exists, remove it before adding new ]-
                                $Worksheet.Cells($Row, $Col).Comment.Delete()
                            }  
                            [void]$WorkSheet.cells.Item($Row, $Col).AddComment("Detected MAC:"+$Data)
                            $Data = $Existing
                        }Else{
                            $Worksheet.Cells($Row, $Col).Font.Bold = $True
                            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 10       #--[ Green, Replace bad MAC with detected ]--
                            If ($Null -ne $Worksheet.Cells($Row, $Col).Comment()){  #--[ If a previous comment exists, remove it before adding new ]-
                                $Worksheet.Cells($Row, $Col).Comment.Delete()
                            }  
                            $Data = $Obj.RealMAC
                        }
                    }Else{                                                          #--[ An issue exists somewhere ]--
                        $Worksheet.Cells($Row, $Col).Font.Bold = $False
                        $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3            #--[ Red ]--
                        If ($Null -ne $Worksheet.Cells($Row, $Col).Comment()){      #--[ If a previous comment exists, remove it before adding new ]-
                            $Worksheet.Cells($Row, $Col).Comment.Delete()
                        }  
                        [void]$WorkSheet.cells.Item($Row, $Col).AddComment("Detected MAC:"+$Data)
                        $Data = $Existing
                    }
                }
                "url" {  
                    $WorkSheet.Hyperlinks.Add($WorkSheet.Cells.Item($Row, $Col),$NewData.Split(";")[0],"","",$NewData.Split(";")[1]) | Out-Null
                }
                Default {
                    If ($Existing -ne $Data){
                        $Worksheet.Cells($Row, $Col).Font.Bold = $true
                        $Worksheet.Cells($Row, $Col).Font.ColorIndex = 7            #--[ Violet to denote a change ]--
                    }Else{                       
                        If ($Format -eq "red"){
                            $Worksheet.Cells($Row, $Col).Font.Bold = $True
                            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3
                        }ElseIf ($Format -eq "green"){
                            $Worksheet.Cells($Row, $Col).Font.Bold = $True
                            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 10
                        }Else{
                            If (($col -eq 10 -or $Col -eq 12) -and ($existing -eq "false")){    #--[ Keep formatting for FTP and SNMP3 if it has not changed ]--
                                $Worksheet.Cells($Row, $Col).Font.Bold = $True
                                $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3
                            }Else{
                                $Worksheet.Cells($Row, $Col).Font.Bold = $False
                                $Worksheet.Cells($Row, $Col).Font.ColorIndex = 0
                            }
                            If ($Null -ne $Worksheet.Cells($Row, $Col).Comment()){              #--[ If a previous comment exists, remove it ]-
                                $Worksheet.Cells($Row, $Col).Comment.Delete()
                            }                
                        }
                    } 
                }
            }
        }
    } 

    #--[ Pre-Write Cleanup ]-------------------------------------------------------------

    If ($Col -eq "2"){   #--[ Color coding for the hostname cell ]--       
        #--[ This gets processed first so If a previous hostname comment exists, remove it ]-
        If ($Null -ne $Worksheet.Cells($Row, 2).Comment()){  
            $Worksheet.Cells($Row, 2).Comment.Delete()
        }
        If (($Data -ne $Existing) -or ($Format -eq "Red")){   #--[ Color code red and add a comment to denote nslookup failure ]--
            $Worksheet.Cells($Row,2).Font.Bold = $True
            $Worksheet.Cells($Row,2).Font.ColorIndex = 3
            $Comment = "WARNING: Hostname DNS lookup failed! ( "+$Data+" )"
            [void]$WorkSheet.cells.Item($Row, 2).AddComment($Comment) 
        }
        $Data = $Existing       
    }
 
    If (($Existing -eq "Replace") -or ($Data -eq "Replace")){
        $Worksheet.Cells($Row, $Col).Font.Bold = $True
        $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3
    }
    If ($Data -eq "N/A"){
        $Worksheet.Cells($Row, $Col).Font.Bold = $False
        $Worksheet.Cells($Row, $Col).Font.ColorIndex = 0
    }
    If (($Data -eq "existing") -or ($Data -eq "")){                         #--[ If no data use existing ]--
        If ($Existing -eq ""){
            $Data = ""
        }Else{
            $Data = $Existing
            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 0                #--[ Using existing data, set to black ]--
            $Worksheet.Cells($Row, $Col).Font.Bold = $False
        }
    }
 
    If (($Col -eq "25") -and ($Data -ne "Passed")){                         #--[ Latest self test has failed.  Flag the hostname ]--
        Switch ($Data) {
            "Unknown" {
                $AddComment = "WARNING: Self-Test may not be executing!"
                $Worksheet.Cells($Row,2).Font.ColorIndex = 45
            } 
            "Failed" {
                $AddComment = "WARNING: Self-Test has failed!"
                $Worksheet.Cells($Row,2).Font.ColorIndex = 3
            }
            Default {
                $AddComment = "WARNING: Self-Test status is unknown."
                $Worksheet.Cells($Row,2).Font.ColorIndex = 45
            } 
        }
        If ($Null -ne $Worksheet.Cells($Row, 2).Comment()){ 
            $Comment = $Worksheet.Cells.item($Row, 2).comment.text()        #--[ Copy the existing comment first ]--   
            $Worksheet.Cells($Row, 2).Comment.Delete()                      #--[ Clear the existing comment ]--
            [void]$WorkSheet.cells.Item($Row, 2).AddComment($Comment+"`n`n"+$AddComment)   #--[ Re-add comments ]--
        }Else{
            [void]$WorkSheet.cells.Item($Row, 2).AddComment($AddComment)    #--[ Add the new comment only ]--
        }
        If ($Comment -like "*DNS*"){
            $Worksheet.Cells($Row,2).Font.ColorIndex = 3                    #--[ Put the color back if comment was DNS ]--
        }

    }

    If (($Col -eq "28") -and ($Data -ne "N/A")){                            #--[ Colorize Hostname orange if runtime is under 30 minutes ]--
        $Hr = $Data.Split(" ")[0]
        $Min = $Data.Split(" ")[2]
        If (([Int]$Hr -lt 1) -and ([Int]$Min -lt 30)){
            $Worksheet.Cells($Row,2).Font.Bold = $True
            $Worksheet.Cells($Row,2).Font.ColorIndex = 45  
            $AddComment = "WARNING: Runtime is under 30 minutes!"
            If ($Null -ne $Worksheet.Cells($Row, 2).Comment()){ 
                $Comment = $Worksheet.Cells.item($Row, 2).comment.text()        #--[ Copy the existing comment first ]--   
                $Worksheet.Cells($Row, 2).Comment.Delete()                      #--[ Clear the existing comment ]--
                [void]$WorkSheet.cells.Item($Row, 2).AddComment($Comment+"`n`n"+$AddComment)   #--[ Re-add comments ]--
            }Else{
                [void]$WorkSheet.cells.Item($Row, 2).AddComment($AddComment)    #--[ Add the new comment only ]--
            }
            If (($Comment -like "*DNS*") -or ($Comment -like "*Failed*")){
                $Worksheet.Cells($Row,2).Font.ColorIndex = 3                    #--[ Put the color back if comment was DNS ]--
            }
        }
    }

    #--[ Write processed data to Excel ]--------------------------------------------------
    If ($Debug){write-host "        Writing :"$Data -ForegroundColor Magenta }
    $ErrorActionPreference = "stop" 
    Try{
        $WorkSheet.cells.Item($Row, $Col) = $Data
    }Catch{
        $WorkSheet.cells.Item($Row, $Col) = $Data.ToString()
        $Msg = "Excel Error: "+$data+"  -  "+$_.Exception.Message
        $WorkSheet.cells.Item($Row, $Col) = $_.Exception.Message
        StatusMsg $Msg "Red"
        $Data.GetType()
    }   
#    Return $Anomaly
}

Function CallPlink ($IP,$command){
    $ErrorActionPreference = "silentlycontinue"  #--[ Must be set as shown or plink will fail ]--
    $Switch = $False
    $UN = $Env:USERNAME
    $DN = $Env:USERDOMAIN
    $UID = $DN+"\"+$UN
    If (Test-Path -Path $Script:PasswordFile){
        $Base64String = (Get-Content $Script:KeyFile)
        $ByteArray = [System.Convert]::FromBase64String($Base64String)
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UID, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $ByteArray)
    }Else{
        $Credential = Get-Credential
    }
    #------------------[ Decrypted Result ]-----------------------------------------
    $Password = $Credential.GetNetworkCredential().Password
    $Domain = $Credential.GetNetworkCredential().Domain
    $Username = $Domain+"\"+$Credential.GetNetworkCredential().UserName

    If (Test-Connection -ComputerName $IP -count 1 -BufferSize 16 -Quiet) {
        #--[ Detect and store SSH key in local registry if needed ]--
        # StatusMsg "Automatically storing SSH key if needed." "Magenta"
        # Write-Output "Y" | 
        # plink-v52.exe -ssh -pw $password $username@$IP #"exit" #*>&1
        # plink-v73.exe -ssh -pw $password $username@$IP #-batch #"exit" #*>&1
        # Start-Sleep -Milliseconds 500
        #------------------------------------------------------------
        StatusMsg "Plink IP: $IP" "Magenta"
        #$test = @(plink-v73.exe -ssh -no-antispoof -pw $Password $username@$IP $command ) #*>&1)
        $test = @(plink-v73.exe -ssh -no-antispoof -batch -pw $Password $Username@$IP $command *>&1)

        If ($test -like "*abandoned*"){
            StatusMsg "Switching Plink version" "Magenta"
            $Switch = $true
        }Else{
            StatusMsg 'Plink version 73 test passed' 'Magenta'
        }
        If ($Switch){
            $Msg = 'Executing Plink v52 (Command = '+$Command+')'
            StatusMsg $Msg 'blue'
            $Result = @(plink-v52.exe -ssh -no-antispoof -batch -pw $Password $username@$IP $command *>&1) 
            $Result
        }Else{
            $ErrorActionPreference = "continue"
            $Msg = 'Executing Plink v73 (Command = '+$Command+')'
            StatusMsg $Msg 'magenta'
                $Result = @(plink-v73.exe -ssh -no-antispoof -batch -pw $Password $username@$IP $command *>&1)
                $Result
        }
        ForEach ($Line in $Result){
            If ($Line -like "*denied*"){
                $Result = "ACCESS-DENIED"
                Break
            } 
        }
        StatusMsg "Command completed..." "Magenta"
        Return $Result
    }Else{
        StatusMsg "Pre-Plink PING check FAILED" "Red"
    }
} 

Function TCPportTest ($Target, $Port, $Debug){
    Try{
        #$Result = Test-NetConnection -ComputerName $Target -Port $Port #-ErrorAction SilentlyContinue -WarningAction SilentlyContinue #-InformationLevel Quiet
        $Result = New-Object System.Net.Sockets.TcpClient($Target, $Port) -ErrorAction:Stop
    }Catch{
        Return $_.Exception.Message
    }
    If ($Debug){
        Write-host "`nFTP Debug :" $Result.connected -foregroundcolor red
    }
    return $Result
}

Function SMNPv3Walk ($Target,$OID,$Debug){
    $WalkRequest = @{
        UserName   = $Script:SMNPv3User
        Target     = $Target
        OID        = $OID
        AuthType   = 'MD5'
        AuthSecret = $Script:SNMPv3Secret
        PrivType   = 'DES'
        PrivSecret = $Script:SNMPv3Secret
        #Context    = ''
    }
    $Result = Invoke-SNMPv3Walk @WalkRequest | Format-Table -AutoSize
    If ($Debug){write-host "SNMpv3 Debug :" $Result }
    Return $Result
}

Function GetSNMPv1 ($Target,$OID,$Debug) {
    $SNMP = New-Object -ComObject olePrn.OleSNMP
    $erroractionpreference = "Stop"
    Try{
        $snmp.open($Target,$Script:SMNPv3User,2,1000)
        $Result = $snmp.get($OID)
    }Catch{
        Return $_.Exception.Message        
    }
    If ($Debug){ write-host "SNMpv1 Debug :" $Result }
    Return $Result
    #$snmp.gettree('.1.3.6.1.2.1.1.1')
}

Function GetSMNPv3 ($Target,$OID,$Debug,$Test){
    If ($Test){  #--[ If 1st user tests positive on 1st use, use it by setting the global variable below ]--
        $GetRequest1 = @{
            UserName   = $Script:SMNPv3User
            Target     = $Target
            OID        = $OID.Split(",")[1]
            AuthType   = 'MD5'
            AuthSecret = $Script:SNMPv3Secret
            PrivType   = 'DES'
            PrivSecret = $Script:SNMPv3Secret
        }
        Try{
            $Result = Invoke-SNMPv3Get @GetRequest1 -ErrorAction:Stop
            If ($Result -like "*Exception*"){
                $Script:v3UserTest = $False  
            }Else{
                $Script:v3UserTest = $True  #--[ Global v3 user variable ]--
            }
        }Catch{
            If ($Debug){
                write-host $_.Exception.Message -ForegroundColor Cyan
                write-host " -- SNMPv3 User 1 failed..." -ForegroundColor red
            }
        }
    }Else{  #--[ User 1 has failed so use user 2 instead ]--
        $GetRequest2 = @{
            UserName   = $Script:SMNPv3AltUser
            Target     = $Target
            OID        = $OID.Split(",")[1]
            #AuthType   = 'MD5'
            #AuthSecret = $Script:SNMPv3Secret
            #PrivType   = 'DES'
            #PrivSecret = $Script:SNMPv3Secret
        }
        Try{
            $Result = Invoke-SNMPv3Get @GetRequest2 -ErrorAction:Stop
        }Catch{
            If ($Result -like "*Exception*"){
                write-host " -- SNMPv3 User 2 failed... No SNMPv3 access..." -ForegroundColor red
                write-host $_.Exception.Message -ForegroundColor Blue
            }
        }
    }
    If ($Debug){
        Write-Host "  -- SNMPv3 Debug: " -ForegroundColor Yellow -NoNewline
        If ($Test){
            Write-Host "SNMP User 2  " -ForegroundColor Green -NoNewline
        }Else{
            Write-Host "SNMP user 1  " -ForegroundColor Green -NoNewline
        }
        Write-Host $OID.Split(",")[0]"  " -ForegroundColor Cyan -NoNewline
        Write-Host $Result    
    }
    Return $Result
}

Function GetMAC ($WorkSheet,$Obj,$Row,$Col,$Debug){
    StatusMsg "Pulling MAC address from network switch..." "Magenta"
    $MAC = ""
    $Existing = $WorkSheet.Cells.Item($Row,29).Text        #\
    $Switch = $WorkSheet.Cells.Item($Row,6).Text           #--[ Read existing spreadsheet cell data for comparison ]-- 
    $Port = $WorkSheet.Cells.Item($Row,7).Text             #/

    If ($Port -NotLike "*Gi*"){  #--[ Assuming all the ports are gig ports, check for spelling issues ]--
        StatusMsg "1 Switchport ID issue detected... Please check spreadsheet..." "red"
    }

    Try{
        $Cmd = "sh mac addr | i "+$Port
        $SwitchIP = [string](nslookup $Switch 10.40.9.11 2>&1)
        $SwitchIP = ($SwitchIP.Split(":")[4]).Trim()
        $SwitchData = CallPlink $SwitchIP $Cmd
    }Catch{
        $MAC = "Not Detected"
    }

    Try{
        $Found = select-string "([A-Za-z0-9]+(\.[A-Za-z0-9]+)+)" -inputobject $SwitchData -ErrorAction:SilentlyContinue
        $MAC = ((($Found.Matches.groups[0].value) -Replace '\.', '') -replace '..(?!$)', '$&:').ToUpper()
    }catch{
        $MAC = "Not Detected"
    }
    
    If ($Debug){
        $Color = "yellow"
        StatusMsg "Switch name = $switch" $Color
        StatusMsg "Switch IP = $SwitchIP" $Color
        StatusMsg "Switch port = $port" $Color
        StatusMsg "Existing MAC = $existing" $Color
        StatusMsg "Detected MAC = $mac" $Color
        StatusMsg "Line returned from switch: $SwitchData" $Color
    }
    
    $Obj | Add-Member -MemberType NoteProperty -Name "SwitchIP" -Value $SwitchIP
    $Obj | Add-Member -MemberType NoteProperty -Name "SwitchName" -Value $Switch
    $Obj | Add-Member -MemberType NoteProperty -Name "SwitchPort" -Value $Port
    $Obj | Add-Member -MemberType NoteProperty -Name "RealMAC" -Value $MAC
    Return $Obj
}

Function OpenExcel ($Excel,$ExcelWorkingCopy,$SheetName,$Console) {
    If (Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy" -PathType Leaf){
        $Script:SpreadSheet = "Existing"
        $WorkBook = $Excel.Workbooks.Open("$PSScriptRoot\$ExcelWorkingCopy")
        $WorkSheet = $Workbook.WorkSheets.Item($SheetName)
        $WorkSheet.activate()
    }Else{
        $Script:SpreadSheet = "New"
        write-warning "new"
        $Workbook = $Excel.Workbooks.Add()
        $Worksheet = $Workbook.Sheets.Item(1)
        $Worksheet.Activate()
        $WorkSheet.Name = "UPS"
        [int]$Col = 1
        $WorkSheet.cells.Item(3,$Col++) = "LAN IP Address"      # A
        $WorkSheet.cells.Item(3,$Col++) = "Host Name"           # B
        $WorkSheet.cells.Item(3,$Col++) = "Facility"            # C
        $WorkSheet.cells.Item(3,$Col++) = "IDF"                 # D
        $WorkSheet.cells.Item(3,$Col++) = "Location"            # E
        $WorkSheet.cells.Item(3,$Col++) = "Switch"              # F
        $WorkSheet.cells.Item(3,$Col++) = "Switch Port"         # G
        $WorkSheet.cells.Item(3,$Col++) = "VLAN"                # H  
        $WorkSheet.cells.Item(3,$Col++) = "Online"              # I
        $WorkSheet.cells.Item(3,$Col++) = "FTP"                 # J
        $WorkSheet.cells.Item(3,$Col++) = "SNMP1"               # K
        $WorkSheet.cells.Item(3,$Col++) = "SNMP3"               # L
        $WorkSheet.cells.Item(3,$Col++) = "DeviceType"          # M   
        $WorkSheet.cells.Item(3,$Col++) = "Manufacturer"        # N               
        $WorkSheet.cells.Item(3,$Col++) = "UPS Model #"         # O
        $WorkSheet.cells.Item(3,$Col++) = "UPS Model Name"      # P
        $WorkSheet.cells.Item(3,$Col++) = "UPS Serial #"        # Q
        $WorkSheet.cells.Item(3,$Col++) = "UPS Firmware Ver"    # R            
        $WorkSheet.cells.Item(3,$Col++) = "UPS Mfg Date"        # S
        $WorkSheet.cells.Item(3,$Col++) = "Age (Yrs)"           # S
        $WorkSheet.cells.Item(3,$Col++) = "EOL Date"            # T
        $WorkSheet.cells.Item(3,$Col++) = "EOS Date"            # U
        $WorkSheet.cells.Item(3,$Col++) = "LDOS Date"           # V
        $WorkSheet.cells.Item(3,$Col++) = "Last Test Date"      # W
        $WorkSheet.cells.Item(3,$Col++) = "Last Test Result"    # X
        $WorkSheet.cells.Item(3,$Col++) = "Battery Pack"        # Y            
        $WorkSheet.cells.Item(3,$Col++) = "Last Batt Change"    # Z
        $WorkSheet.cells.Item(3,$Col++) = "Battery Runtime"     # AA
        $WorkSheet.cells.Item(3,$Col++) = "Replace Batt?"       # AB  
        $WorkSheet.cells.Item(3,$Col++) = "Replace on Date"     # AC 
        $WorkSheet.cells.Item(3,$Col++) = "Battery Load %"      # AD          
        $WorkSheet.cells.Item(3,$Col++) = "NMC Model"           # AE
        $WorkSheet.cells.Item(3,$Col++) = "NMC MAC"             # AF
        $WorkSheet.cells.Item(3,$Col++) = "NMC Serial #"        # AG           
        $WorkSheet.cells.Item(3,$Col++) = "NMC Hardware Ver"    # AH
        $WorkSheet.cells.Item(3,$Col++) = "NMC AOS Ver"         # AI
        $WorkSheet.cells.Item(3,$Col++) = "NMC AOS Firmware"    # AJ
        $WorkSheet.cells.Item(3,$Col++) = "NMC App Ver"         # AK
        $WorkSheet.cells.Item(3,$Col++) = "NMC App Firmware"    # AL
        $WorkSheet.cells.Item(3,$Col++) = "NMC Mfg Date"        # AM
        $WorkSheet.cells.Item(3,$Col++) = "Serviced By"         # AN
        $WorkSheet.cells.Item(3,$Col++) = "Comments"            # AO
        $WorkSheet.cells.Item(3,$Col++) = "URL"                 # AP
        $WorkSheet.cells.Item(3,$Col++) = "Date Inspected"      # AQ
        $Range = $WorkSheet.Range(("A3"),("AQ3")) 
        $Range.font.bold = $True
        $Range.HorizontalAlignment = -4108  #Alignment Middle
        $Range.Font.ColorIndex = 44
        $Range.Font.Size = 12
        $Range.Interior.ColorIndex = 56
        $Range.font.bold = $True
        1..4 | ForEach {
            $Range.Borders.Item($_).LineStyle = 1
            $Range.Borders.Item($_).Weight = 4
        }
        $Resize = $WorkSheet.UsedRange
        [Void]$Resize.EntireColumn.AutoFit()
    }
    Return $WorkBook
}

Function GetSource ($SourcePath,$ExcelSourceFile,$ExcelWorkingCopy){
    StatusMsg "Excel working copy was not found, copying from source..." "Magenta"
    If (Test-Path -Path "$SourcePath\$ExcelSourceFile" -PathType Leaf){
        Try{
            Copy-Item -Path "$SourcePath\$ExcelSourceFile"  -Destination "$PSScriptRoot\$ExcelWorkingCopy" -force -ErrorAction:Stop
            Return $True
        }Catch{
            write-host $_.Exception.Message
            write-host $_.Error.Message
            Return $False   
            StatusMsg "Copy failed... " "red" 
        }
    }Else{   
        StatusMsg "Source file check failed... " "red"
        Return $False
    }
}

#--[ End of Functions ]-------------------------------------------------------

$TransferTable = @{
    "1" = "No events"
    "2" = "High line voltage"
    "3" = "Brownout"
    "4" = "Loss of mains power"
    "5" = "Small temporary power drop"
    "6" = "Large temporary power drop"
    "7" = "Small spike"
    "8" = "Large spike"
    "9" = "UPS self test"
    "10" = "Excessive input voltage fluctuation"
}

$OIDArray = @()
$OIDArray += ,@('LastTestResult','.1.3.6.1.4.1.318.1.1.1.7.2.3.0') #--[ 1=passed, 2=failed, 3=never run ]--
$OIDArray += ,@('LastTestDate','.1.3.6.1.4.1.318.1.1.1.7.2.4.0')  #--[ returns a date or nothing ]--
$OIDArray += ,@('UPSSerial','.1.3.6.1.4.1.318.1.1.1.1.2.3.0')
#$OIDArray += ,@('UPSModelName','.1.3.6.1.4.1.318.1.1.1.1.1.1.0')
$OIDArray += ,@('UPSModelName','.1.3.6.1.4.1.318.1.4.2.2.1.5.1')
$OIDArray += ,@('UPSModelNum','.1.3.6.1.4.1.318.1.1.1.1.2.5.0')
$OIDArray += ,@('UPSMfgDate','.1.3.6.1.4.1.318.1.1.1.1.2.2.0')
#--[ MfgDate from SN:  xx1915xxxxxx means mfg in 2019, 15th week.  ]--
$OIDArray += ,@('UPSIDName','.1.3.6.1.2.1.33.1.1.5.0')
$OIDArray += ,@('FirmwareVer','.1.3.6.1.4.1.318.1.1.1.1.2.1.0')
$OIDArray += ,@('Mfg','.1.3.6.1.2.1.33.1.1.1.0')
$OIDArray += ,@('MfgDate','.1.3.6.1.4.1.318.1.1.1.1.2.2.0')
$OIDArray += ,@('MAC','.1.3.6.1.2.1.2.2.1.6.2')
$OIDArray += ,@('Location','.1.3.6.1.2.1.1.6.0')
$OIDArray += ,@('Contact','.1.3.6.1.2.1.1.4.0')       
$OIDArray += ,@('HostName','.1.3.6.1.2.1.1.5.0')       
$OIDArray += ,@('NMC','.1.3.6.1.2.1.1.1.0')   
$OIDArray += ,@('BattFreqOut','.1.3.6.1.4.1.318.1.1.1.4.2.2.0')
$OIDArray += ,@('BattVOut','.1.3.6.1.4.1.318.1.1.1.4.2.1.0')
$OIDArray += ,@('BattVIn','.1.3.6.1.4.1.318.1.1.1.3.2.1.0')
$OIDArray += ,@('BattFreqIn','.1.3.6.1.4.1.318.1.1.1.3.2.4.0')
$OIDArray += ,@('BattActualV','.1.3.6.1.4.1.318.1.1.1.2.2.8.0')
$OIDArray += ,@('BattCurrentAmps','.1.3.6.1.4.1.318.1.1.1.2.2.9.0')
$OIDArray += ,@('BattChangedDate','.1.3.6.1.4.1.318.1.1.1.2.1.3.0')
$OIDArray += ,@('BattCapLeft','.1.3.6.1.4.1.318.1.1.1.2.2.1.0')
$OIDArray += ,@('BattRunTime','.1.3.6.1.4.1.318.1.1.1.2.2.3.0')
$OIDArray += ,@('BattReplace','.1.3.6.1.4.1.318.1.1.1.2.2.4.0')
$OIDArray += ,@('BattReplaceDate','.1.3.6.1.4.1.318.1.1.1.2.2.21.0')
$OIDArray += ,@('BattSKU','.1.3.6.1.4.1.318.1.1.1.2.2.19.0')
$OIDArray += ,@('BattExtSKU','.1.3.6.1.4.1.318.1.1.1.2.2.20.0')
$OIDArray += ,@('BattTemp','.1.3.6.1.4.1.318.1.1.1.2.2.2.0')
$OIDArray += ,@('ACVIn','.1.3.6.1.4.1.318.1.1.1.3.2.1.0')
$OIDArray += ,@('ACFreqIn','.1.3.6.1.4.1.318.1.1.1.3.2.4.0')
$OIDArray += ,@('LastXfer','.1.3.6.1.4.1.318.1.1.1.3.2.5.0')
$OIDArray += ,@('UPSVOut','.1.3.6.1.4.1.318.1.1.1.4.2.1.0')
$OIDArray += ,@('UPSFreqOut','.1.3.6.1.4.1.318.1.1.1.4.2.2.0')
$OIDArray += ,@('UPSOutLoad','.1.3.6.1.4.1.318.1.1.1.4.2.3.0')
$OIDArray += ,@('UPSOutAmps','.1.3.6.1.4.1.318.1.1.1.4.2.4.0')    

$RBCTable = @{
    #--[ Unfortunately APC does not provide full model numbers via SNMP except on the most very newest ]--
    #--[ UPS devices.  That makes determining the RBC very hard.  The RBC changes depening on model ]--
    #--[ options even though the name of the UPS covers multiple models as shown below. ]--
    'SC1500' = 'RBC59'
    'SC420' = 'RBC2'
    'SC750' = 'RBC2'
    'SMT1500' = 'RBC7'           #--[ also RBC139 and RBC49 ]--
    'SMT1500C' = 'RBC7'          #--[ also RBC139 and RBC49 ]--
    'SMT1500RM2U' = 'RBC133'     #--[ Also RBC116 ]--
    'SMT1500RM2UC' = 'RBC159'
    'SMT750C' = 'RBC48'
    'SRT3000RMXLT' = 'RBC152'
    'SU1400R2BX120' = 'RBC24'
    'SU1400R2BX135' = 'RBC24'
    'SU1400RM2U' = 'RBC24'
    'SUA1000' = 'RBC6'
    'SUA1500 ' = 'RBC7'
    'SUA1500RM2U' = 'RBC24'
    'SUA3000RM2U' = 'RBC43'
    'SUA2200RM2U' = 'RBC43'
    'SRT3000RMXL2U' = 'RBC152'
    'SUA5000RMT5U' = 'RBC55'
    <#--[ Partial list of UPS model #, Name, & RBC ]--
    SC1500	        Smart-UPS 1500	    RBC59		
    SC420	        SC420	            RBC2		
    SC750	        Smart-UPS 750   	RBC2		
    SMT1500C	    Smart-UPS 1500  	RBC7	RBC139	RBC49
    SMT1500RM2U	    Smart-UPS 1500	    RBC133	RBC116	
    SMT1500RM2U	    Smart-UPS 1500 RM	RBC133	RBC116	
    SMT1500RM2UC	Smart-UPS 1500	    RBC159		
    SMT1500RM2UC	Smart-UPS 1500	    RBC159		
    SMT750C	        Smart-UPS 750	    RBC48		
    SMT750C	        Smart-UPS 750	    RBC48		
    SRT3000RMXLT	Smart-UPS 2200 RM	RBC152		
    SRT3000RMXLT	Smart-UPS SRT 3000	RBC152		
    SU1400R2BX120	Smart-UPS 1400 RM	RBC24		
    SU1400R2BX135	Smart-UPS 1400 RM	RBC24		
    SU1400RM2U	    Smart-UPS 1400 RM	RBC24		
    SU1400RM2U	    Smart-UPS 1400 RM	RBC24		
    SUA1000	        Smart-UPS 1000	    RBC6		
    SUA1400RM2U	    Smart-UPS 1400 RM	RBC24		
    SUA1500 	    Smart-UPS 1500	    RBC7	RBC139	RBC49
    SUA1500RM2U	    Smart-UPS 1500	    RBC133		
    SUA1500RM2U	    Smart-UPS 1500 RM	RBC24		
    SUA2200RM2U	    Smart-UPS 2200 RM	RBC43		
    SUA2200RM2U	    Smart-UPS 2200 RM	RBC43		
    SUA3000RM2U	    Smart-UPS 3000 RM	RBC43		
    SUA5000RMT5U	Smart-UPS 1500 RM	RBC55		
    SUA5000RMT5U	Smart-UPS 5000	    RBC55		
    SUM3000RMXL2U	Smart-UPS 3000 XLM	RBC43		
    #>
}


#==[ Begin ]==============================================================

LoadConfig
StatusMsg "Processing UPS Devices" "Yellow"
$erroractionpreference = "stop"

#--[ Close copies of Excel that PowerShell has open ]--
If ($CloseOpen){
    $ProcID = Get-CimInstance Win32_Process | where {$_.name -like "*excel*"}
    ForEach ($ID in $ProcID){  #--[ Kill any open instances to avoid issues ]--
        Foreach ($Proc in (get-process -id $id.ProcessId)){
            if (($ID.CommandLine -like "*/automation -Embedding") -Or ($proc.MainWindowTitle -like "$ExcelWorkingCopy*")){
                Stop-Process -ID $ID.ProcessId -Force
                StatusMsg "Killing any existing open PowerShell instance of Excel..." "Red"
                Start-Sleep -Milliseconds 100
            }
        }
    }
}

#--[ Create new Excel COM object ]--
$Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
StatusMsg "Creating new Excel COM object..." "Magenta"

#--[ Make a backup of the working copy, keep only the last 10 ]--
If (($SafeUpdate)-And (Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy")){
    StatusMsg "Safe-Update Enabled. Creating a backup copy of the working spreadsheet..." "Green"
    $Backup = $DateTime+"_"+$ExcelWorkingCopy+".bak"
    Copy-Item -Path "$PSScriptRoot\$ExcelWorkingCopy"  -Destination "$PSScriptRoot\$Backup"
    #--[ Only keep 10 of the last backups ]-- 
    Get-ChildItem -Path $PSScriptRoot | Where-Object {(-not $_.PsIsContainer) -and ($_.Name -like "*.bak")} | Sort-Object -Descending -Property LastTimeWrite | Select-Object -Skip 10 | Remove-Item
}

#--[ If this file exists the IP list will be pulled from it ]--
If (Test-Path -Path $TestFileName){   
    $ListFileName = $TestFileName   #--[ Select an alternate short IP text file to use ]--
}Else{ 
    $ListFileName = "$PSScriptRoot\$IPTextFile"   #--[ Select the normal IP text file to use ]--
}

#--[ Identify IP address list source and process. ]--
If (Test-Path -Path $ListFileName){  #--[ If text file exists pull from there. ]--
    $IPList = Get-Content $ListFileName          
    StatusMsg "IP text list was found, loading IP list from it... " "green" 
    If (Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy"){
        StatusMsg ">>>     WARNING: Working copy already exists.     <<<" "Yellow"
        StatusMsg ">>>  New copy will be created and NOT over-write. <<<" "Yellow"
        StatusMsg ">>> Remember to delete IP file prior to next run. <<<" "Yellow"
        Start-Sleep -Seconds 5
    }
    StatusMsg "Creating new Spreadsheet..." "green"  
    $Excel.Visible = $True
    $WorkBook = OpenExcel $Excel $ExcelWorkingCopy "UPS" $Console #--[ Create a new spreadsheet.  Default option. ]--
    $WorkSheet = $Workbook.WorkSheets.Item("UPS")
    $WorkSheet.activate()
}else{ 
    StatusMsg "IP text list not found, Attempting to process spreadsheet... " "cyan"
    If (Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy" -PathType Leaf){
        $Excel.Visible = $True
        $WorkBook = OpenExcel $Excel $ExcelWorkingCopy "UPS" $Console #--[ Open the existing spreadsheet if detected. ]--            
    }Else{
        If (GetSource $SourcePath $ExcelSourceFile $ExcelWorkingCopy){
            StatusMsg "Copied new working copy of source to script folder..." "Green"
            $Excel.Visible = $True
            $WorkBook = OpenExcel $Excel $ExcelWorkingCopy "UPS" $Console #--[ Open the existing spreadsheet if detected. ]-- 
            StatusMsg "Removing un-needed worksheets..." "Green"
            $Excel.displayalerts = $False
            ForEach ($WorkSheet in $WorkBook.Worksheets){
                If ($WorkSheet.Name -ne "UPS"){
                    $WorkSheet.Delete()
                }
            }         
        }Else{
            StatusMsg "Existing spreadsheet not found, Source copy failed, Nothing to process.  Exiting... " "red"
            Break;break
        }
    }
    $WorkSheet = $Workbook.WorkSheets.Item("UPS")
    $WorkSheet.activate()    
    $Row = 4   
    $IPList = @() 
    Do {
        $IPList += ,@($Row,$WorkSheet.Cells.Item($row,1).Text)
        $Row++
    } Until (
        $WorkSheet.Cells.Item($row,1).Text -eq ""   #--[ Condition that stops the loop if it returns true ]--
    )
}

$Excel.DisplayAlerts = $false
$Workbook.SaveAs("$PSScriptRoot\$ExcelWorkingCopy") 
StatusMsg "Saving workbook" "Magenta"

ForEach ($Target in $IPList){  #--[ Are we pulling from Excel or a text file?  Jagged has row numbers from Excel ]--
    if ($Target.length -eq 2){
        $Jagged = $True
    }Else{
        $Jagged = $False
    }
}

#==[ Process items collected in $IPList, from spreadsheet, or text file as appropriate ]===============================
$Row = 4   
ForEach ($Target in $IPList){
    If ($Jagged){
        $Row = $Target[0]
        $Target = $Target[1]
    }

    If ($Console){Write-Host "`nCurrent Target  :"$Target"  (Row:"$Row")" -ForegroundColor Yellow }

    #--[ Sets row color to pale blue to denote which is being worked on ]--
    $Excel.ActiveSheet.UsedRange.Rows.Item($Row-1).Interior.ColorIndex = 20  
    #----------------------------------------------------------------------
   
    $Obj = New-Object -TypeName psobject   #--[ Collection for Results ]--
    Try{
        $HostLookup = (nslookup $Target $Script:DNS 2>&1) 
        $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value ($HostLookup[3].split(":")[1].TrimStart()).Split(".")[0] -force
        $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $True
    }Catch{
        $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value "Not Found" -force
        $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $False
    }
    
    $Obj | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $Target -force
    
    #--[ Validate and clean up NMC MAC address by polling assigned switch ]--
    $Obj = GetMAC $WorkSheet $Obj $Row $Col $Debug

    If (Test-Connection -ComputerName $Target -count 1 -BufferSize 16 -Quiet){  #--[ Ping Test ]--
        $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "Online" -force
        StatusMsg "Polling SNMP..." "Magenta"
        If ((!($Debug)) -and ($Console)){Write-Host "  Working." -NoNewline}

        #--[ Test for FTP ]--------------------------------------------------
        $Test = TCPportTest $Target "21" $Debug                
        if ($Test.Connected -like "*True*"){
            $PortTest = "True"
        }Else{
            $PortTest = "False"
        }
        $Obj | Add-Member -MemberType NoteProperty -Name "FTP" -Value $PortTest -force
        If (!($Debug)){Write-Host "." -NoNewline}

        #--[ Test for SNMPv3.  Make sure to include leading comma  ]---------
        $Test = GetSMNPv3 $Target ",1.3.6.1.2.1.1.8.0" $Debug $Script:v3UserTest
        if ($Test -like "*TimeTicks*"){
            $PortTest = "True"
        }Else{
            $PortTest = "False"
        }
        $Obj | Add-Member -MemberType NoteProperty -Name "SNMPv3" -Value $PortTest -force
        If (!($Debug)){Write-Host "." -NoNewline}

        #--[ Test for SNMPv1 ]------------------------------------------------
        $Test = GetSNMPv1 $Target "1.3.6.1.2.1.1.8.0" $Debug
        if ($Test -like "*TimeTicks*"){
            $PortTest = "True"
        }Else{
            $PortTest = "False"
        }
        $Obj | Add-Member -MemberType NoteProperty -Name "SNMPv1" -Value $PortTest -force
    }

    #--[ Only process OIDs if online PLUS SNMPv3 is good ]--------------------------
    If (($Obj.Connection -eq "Online") -and ($Obj.SNMPv3 -ne "False")){  
        ForEach ($Item in $OIDArray){            
            $Result = GetSMNPv3 $Target $Item $Debug $Script:v3UserTest
            If ($Debug){
                Write-Host ' '$Item[0]'='$Result -ForegroundColor yellow
            }Else{
                If ($Console){Write-Host "." -NoNewline}   #--[ Writes a string of dots to show operation is proceeding ]--
            }

            If ($Obj.HostName -like "*chill*" ){
                $Obj | Add-Member -MemberType NoteProperty -Name "DeviceType" -Value "Chiller" -force
                $Obj | Add-Member -MemberType NoteProperty -Name "BattSKU" -Value "N/A" -force
                $Obj | Add-Member -MemberType NoteProperty -Name "BattRunTime" -Value "N/A" -force
                $Obj | Add-Member -MemberType NoteProperty -Name "BattReplaceDate" -Value "N/A" -force
                $Obj | Add-Member -MemberType NoteProperty -Name "BattReplace" -Value "N/A" -force
                $Obj | Add-Member -MemberType NoteProperty -Name "LastXfer" -Value "N/A" -force
                $Obj | Add-Member -MemberType NoteProperty -Name "UPSOutLoad" -Value "N/A" -force
            }Else{
                $Obj | Add-Member -MemberType NoteProperty -Name "DeviceType" -Value "UPS" -force
            }

            #--[ Clean Up Results ]-------------------------------------------------
            Switch ($Item[0]) {
                "MAC"{   #--[ Extract and format MAC Address ]--   
                    $MAC = [System.Text.Encoding]::Default.GetBytes($Result.Value) | ForEach-Object {
                        $_.ToString('X2')
                    }     
                    $SaveVal = $MAC -join ':' 
                    $Obj | Add-Member -MemberType NoteProperty -Name "MAC" -Value $SaveVal -force
                } #>
                "LastXfer" {   #--[ Extract last transfer date ]-- 
                    $SaveVal = $TransferTable[$Result.Value.ToString()]
                    $Obj | Add-Member -MemberType NoteProperty -Name "LastXfer" -Value $SaveVal -force
                } #>
                "HostName" {   #--[ Extract and compare hostname ]--   
                    If ($Obj.Hostname -match $Result.Value.ToString()){
                        $SaveVal = $Result.Value.ToString()                  #--[ Hostnames match ]--
                    }Else{
                        If ($Obj.Hostname -like "*40-UPS*"){
                            $SaveVal = $Obj.Hostname                         #--[ DNS value is good ]--    
                        }ELse{
                            $SaveVal = $Result.Value.ToString()              #--[ DNS wrong, use SNMP ]--
                            $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $False -force
                        }
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "HostName" -Value $SaveVal -force
                } #>
                "LastTestResult" {   #--[ Extract last test result ]--  
                    Switch ($Result.Value){
                        "1" {
                            $SaveVal = "Passed"    
                        }
                        "2" {
                            $SaveVal = "Failed"
                        }
                        "3" {
                            $SaveVal = "Unknown"
                        }
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "LastTestResult" -Value $SaveVal -force
                } #>
                "BattReplace" {   #--[ Extract battery needs replacement flag ]--
                    If ($Obj.DeviceType -eq "UPS"){
                        If ($Result.Value -eq 2){
                            $SaveVal = "Replace"
                        }Else{
                            $SaveVal = "N/A"
                        }   
                    }Else{
                        $SaveVal = "N/A"
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "BattReplace" -Value $SaveVal -force
                } #>
                "BattReplaceDate" {   #--[ Clean up battery replacement date ]--
                     #--[ NOTE: APC seems to recommend that batteries are replaced at least every 5.5 years (66 months)]--
                     If ($Obj.HostName -Like "*chill*"){
                        $SaveVal = "N/A"
                        $Obj | Add-Member -MemberType NoteProperty -Name "BattChangedDate" -Value $SaveVal -force
                    }Else{
                        $DateToday = get-date -Format MM/dd/yyyy
                        $ReplaceDate = (Get-Date -Date $obj.BattChangedDate).AddMonths(66) 
                        $ReplaceDate = Get-Date -Format MM/dd/yyyy -Date $ReplaceDate
                        If (([String]$Result.Value -like "*NoSuch*") -or ($Result.Value -eq "")){  
                            #--[ If the replace-on-date value is blank force it to 5.5 years past last changed ]--                     
                            $SaveVal = $ReplaceDate 
                        }  
                        If((Get-Date -Date $DateToday) -ge (Get-Date -Date $ReplaceDate)){
                            $Obj | Add-Member -MemberType NoteProperty -Name "BattReplace" -Value "Replace" -force
                            $SaveVal = $DateToday #Get-Date -Format MM/dd/yyyy
                        }Else{                
                            $Obj | Add-Member -MemberType NoteProperty -Name "BattReplace" -Value "N/A" -force
                            $SaveVal = $ReplaceDate 
                        }
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "BattReplaceDate" -Value $SaveVal -force
                }
                "Location" {   #--[ Location field on device must be formatted as "facility;IDF;address" separated by a semicolon ]--
                    $SaveVal = $Result.Value.ToString()
                    $Obj | Add-Member -MemberType NoteProperty -Name "Facility" -Value $SaveVal.Split(";")[0] -force
                    $Obj | Add-Member -MemberType NoteProperty -Name "IDF" -Value $SaveVal.Split(";")[1] -force
                    $Obj | Add-Member -MemberType NoteProperty -Name "Location" -Value $SaveVal.Split(";")[2] -force
                } #>
                "NMC" {
                    $Flag = $False
                    $NMCArray = ($Result.Value).ToString().Split(" ")
                    ForEach ($Item in $NMCArray){
                        If ($Flag){
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCSerial" -Value $Item -force
                            $Flag = $False
                        }
                        If ($Item -like "MN:*"){
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCModelNum" -Value $Item.Split(":")[1] -force
                        }
                        If ($Item -like "MD:*"){
                            $String = ($Item.Split(":")[1]).Substring(0,($Item.Split(":")[1]).Length-1)
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCMfgDate" -Value $String -force
                        }
                        If ($Item -like "HR:*"){        #--[ Hardware Revision ]--                    
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCHardwareVer" -Value $Item.Split(":")[1] -force
                        }
                        If ($Item -like "PF:*"){        #--[ AOS Version ]--
                            #$String = ($Item.Split(":")[1]).Substring(0,($Item.Split(":")[1]).Length-1)                           
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCAOSVer" -Value $Item.Split(":")[1] -force
                            #$Obj | Add-Member -MemberType NoteProperty -Name "NMCAOSVer" -Value $String -force
                        }
                        If ($Item -like "AF1:*"){       #--[ Application Version ]--
                            #$String = ($Item.Split(":")[1]).Substring(0,($Item.Split(":")[1]).Length-1)
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCAppVer" -Value $Item.Split(":")[1] -force
                            #$Obj | Add-Member -MemberType NoteProperty -Name "NMCAppVer" -Value $String -force
                        }
                        If ($Item -like "PN:*"){        #--[ AOS Firmware File Version ]--
                            #$String = ($Item.Split(":")[1]).Substring(0,($Item.Split(":")[1]).Length-1)
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCAOSFirmware" -Value $Item.Split(":")[1] -force
                            #$Obj | Add-Member -MemberType NoteProperty -Name "NMCAOSFirmware" -Value $String -force
                        }
                        If ($Item -like "AN1:*"){       #--[ Application Firmware File Version ]--
                            #$String = ($Item.Split(":")[1]).Substring(0,($Item.Split(":")[1]).Length-1)
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCAppFirmware" -Value $Item.Split(":")[1] -force
                            #$Obj | Add-Member -MemberType NoteProperty -Name "NMCAppFirmware" -Value $String -force
                        }
                        If ($Item -like "SN:*"){$Flag = $True}
                    }
                } #>
                "BattRunTime" {
                    If ($Obj.HostName -Like "*chill*"){
                        $SaveVal = "N/A"
                    }Else{
                        $Result = $Result.Value.ToString()
                        $RunHours = $Result.Split(":")[0]
                        $RunMins = $Result.Split(":")[1]
                        #$RunSecs = $Result.Split(":")[2]
                        $SaveVal = $RunHours+" Hrs "+$RunMins+" Min"
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "BattRunTime" -Value $SaveVal -force
                }
                "BattSKU" {
                    If ($Obj.HostName -Like "*chill*"){
                        $SaveVal = "N/A"
                    }Else{
                        $SaveVal = $Result.Value
                        If ($SaveVal -like "*APC*"){                  
                            $Bytes = [System.Text.Encoding]::Unicode.GetBytes($Result.Value)
                            $SaveVal = "" 
                            ForEach ($Value in $Bytes){
                                If ($Value -ne " "){
                                    $SaveVal += ([System.Text.Encoding]::ASCII.GetString($Value)).trim()                
                                }
                            }                    
                        }
                        If ($SaveVal -like "APC*"){
                            $SaveVal = $SaveVal.Substring(3)
                        }
                    }
                    If ($SaveVal -like "*nosuch*"){
                        $SaveVal = ""
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "BattSKU" -Value $SaveVal.ToString() -force
                }
                Default {   #--[ Use values pulled from SNMP for all others ]--
                    If (($Result.Value -eq "") -or ($Null -eq $Result.Value)){
                        $SaveVal = "existing"
                    }Else{
                        $SaveVal = $Result.Value.ToString()
                    }
                    If ($SaveVal -like "NoSuch*"){
                        $SaveVal = ""
                    }ElseIf ($Item -like "*date*"){    #--[ Set dates to a uniform format ]--
                        If ($SaveVal -eq ""){
                            $SaveVal = "existing"
                        }Else{
                            Try {
                                $SaveVal = Get-Date $Result.Value.ToString() -Format MM/dd/yyyy -ErrorAction SilentlyContinue
                            }Catch{ 
                                $SaveVal = $Result.Value.ToString()                   
                            }   
                        }
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name $Item[0] -Value $SaveVal -force
                }   #>         
            }
    
            #--[ Adjustments ]------------------------
            If (($Obj.UPSModelName -like "*Smart-UPS*") -or ($Obj.UPSModelName -like "*Symmetra*")){   #-[ Since most don't respond with Mod # & Mfg, fake it ]--
                $Obj | Add-Member -MemberType NoteProperty -Name "Mfg" -Value "APC" -force
            }            
        }

        #--[ Support Dates ]-------------------------------------------
        #--[ Per the APC web site "APC product warranties are based on a fixed period of time starting from the purchase ]--
        #--[ date of the unit. If proof of purchase cannot be obtained, the warranty duration is calculated from the     ]--
        #--[ manufacture date of the unit. As multiple factors go into determining the validity of a warranty, you must  ]--
        #--[ speak with APC Customer Care Center to obtain the most accurate information.  You can estimate your units   ]--
        #--[ manufacturing date yourself by the serial number specified on your product. The third and fourth digits     ]--
        #--[ indicate the year your unit was manufactured, while the following 2 characters are the week your unit was   ]--
        #--[ manufactured."                                                                                              ]--
        #--[ This basically means nothing.  So, dates are "assumed" here, EOS=unknown, EOL = Mfg+5, and LDOS=Mfg+10      ]--
        #--[ unless it's been predetermined and added to the spreadsheet.                                                ]--

        #--[ Get UPS Age by MFG Date ]--
 
        If ($Obj.UPSMfgDate -eq ""){
            $Obj | Add-Member -MemberType NoteProperty -Name "UPSAge" -Value "existing" -force
        }Else{
            $MfgYear = $Obj.UPSSerial.Substring(2,3)
            $MfgWeek = $Obj.UPSSerial.Substring(4,2)
            $Obj | Add-Member -MemberType NoteProperty -Name "UPSAge" -Value ([int]((New-TimeSpan -Start ([datetime]$Obj.UPSMfgDate) -End $Today).days/365)) -force         
        } 

        $Obj | Add-Member -MemberType NoteProperty -Name "EOLDate" -Value "existing" -force
        $Obj | Add-Member -MemberType NoteProperty -Name "EOSDate" -Value "existing" -force
        $Obj | Add-Member -MemberType NoteProperty -Name "LDOSDate" -Value "existing" -force

    }

    If ($Console){
        Write-host " "
        $Obj
    }    

    #--[ Write each column's data for current row ]--
    StatusMsg "Writing Data to Excel... " "Magenta"
    $Col=1
    $WorkSheet.cells.Item($Row,($Col++)) = $Obj.IPAddress                     # A
    If ($Obj.HostNameLookup){
        Write2Excel $WorkSheet $Row ($Col++) $Obj.HostName "green"             # B  #--[ Apply green emphasis if nslookup succeeds ]-- 
    }Else{
        Write2Excel $WorkSheet $Row ($Col++) $Obj.HostName "red"               # B  #--[ Apply red emphasis if nslookup fails ]--            
    }
    If ($Obj.Connection -eq "Online"){       
        Write2Excel $WorkSheet $Row 9 $Obj.Connection "green"           # I  #--[ Apply green emphasis if target is online ]--
        If ($Obj.SNMPv3 -ne "False"){   
            Write2Excel $WorkSheet $Row ($Col++) $Obj.Facility                 # C    
            Write2Excel $WorkSheet $Row ($Col++) $Obj.IDF                      # D
            Write2Excel $WorkSheet $Row ($Col++) $Obj.Location                 # E    
              #--[ Skip the next cell ]--
              $Col++
            #$WorkSheet.cells.Item($Row, $Col++) = "Switch"             # F    
              #--[ Skip the next cell ]--
              $Col++
            #$WorkSheet.cells.Item($Row, $Col++) = "Switch Port"        # G
              #--[ Skip the next cell ]--
              $Col++
            #$WorkSheet.cells.Item($Row, $Col++) = "VLAN"               # H
              #--[ Skip the next cell.  This one was set above. ]--
              $Col++
            #$WorkSheet.cells.Item($Row, $Col++) = "Online"             # I
            If ($Obj.FTP -eq "False"){
                Write2Excel $WorkSheet $Row ($Col++) $Obj.FTP "red"           # J  #--[ Apply red emphasis if FTP is not enabled ]--
            }Else{
                Write2Excel $WorkSheet $Row ($Col++) $Obj.FTP "green"         # J (Col 10)
            }
            Write2Excel $WorkSheet $Row ($Col++) $Obj.SNMPv1                  # K  #--[ Apply green emphasis if SMNPv3 is enabled ]--
            Write2Excel $WorkSheet $Row ($Col++) $Obj.SNMPv3 "green"          # L
            Write2Excel $WorkSheet $Row ($Col++) $Obj.Mfg                     # M  
            Write2Excel $WorkSheet $Row ($Col++) $Obj.DeviceType              # N   
            Write2Excel $WorkSheet $Row ($Col++) $Obj.UPSModelNum "existing"  # O (Col 15)
            Write2Excel $WorkSheet $Row ($Col++) $Obj.UPSModelName            # P
            Write2Excel $WorkSheet $Row ($Col++) $Obj.UPSSerial               # Q
            Write2Excel $WorkSheet $Row ($Col++) $Obj.FirmwareVer             # R
            Write2Excel $WorkSheet $Row ($Col++) $Obj.MfgDate "date"          # S
            Write2Excel $WorkSheet $Row ($Col++) $Obj.UPSAge                  # T (Col 20)
            Write2Excel $WorkSheet $Row ($Col++) $Obj.EOLDate "date"          # U
            Write2Excel $WorkSheet $Row ($Col++) $Obj.EOSDate "date"          # V
            Write2Excel $WorkSheet $Row ($Col++) $Obj.LDOSDate "date"         # W
            Write2Excel $WorkSheet $Row ($Col++) $Obj.LastTestDate "date"     # X
            Write2Excel $WorkSheet $Row ($Col++) $Obj.LastTestResult          # Y (Col 25)  #--[ Edit write function if this column # changes ]--
            Write2Excel $WorkSheet $Row ($Col++) $Obj.BattSKU "existing"      # Z   #--[ Edit write function if this column # changes ]-- 
            Write2Excel $WorkSheet $Row ($Col++) $Obj.BattChangedDate "date"  # AA 
            Write2Excel $WorkSheet $Row ($Col++) $Obj.BattRunTime "run"       # AB  #--[ Edit write function if this column # changes ]--
            Write2Excel $WorkSheet $Row ($Col++) $Obj.BattReplace             # AC
            Write2Excel $WorkSheet $Row ($Col++) $Obj.BattReplaceDate "date"  # AD  #--[ Edit write function if this column # changes ]--
            Write2Excel $WorkSheet $Row ($Col++) $Obj.UPSOutLoad "run"        # AE (Col 31)    
            Write2Excel $WorkSheet $Row ($Col++) $Obj.NMCModelNum             # AF
            Write2Excel $WorkSheet $Row ($Col++) $Obj.MAC "mac"               # AG
            Write2Excel $WorkSheet $Row ($Col++) $Obj.NMCSerial               # AH
            Write2Excel $WorkSheet $Row ($Col++) $Obj.NMCHardwareVer          # AI (Col 35)
            Write2Excel $WorkSheet $Row ($Col++) $Obj.NMCAOSVer               # AJ
            Write2Excel $WorkSheet $Row ($Col++) $Obj.NMCAOSFirmware          # AK
            Write2Excel $WorkSheet $Row ($Col++) $Obj.NMCAppVer               # AL
            Write2Excel $WorkSheet $Row ($Col++) $Obj.NMCAppFirmware          # AM
            Write2Excel $WorkSheet $Row ($Col++) $Obj.NMCMfgDate "date"       # AN (Col 40)
              #--[ Skip the next cell ]--
            #Write2Excel $WorkSheet $Row 37 "Serviced By"               # AO
            $Col++
              #--[ Skip the next cell ]--
            #Write2Excel $WorkSheet $Row 38 "Comments"                  # AP
            $Col++
              #--[ Skip the next cell ]--
            #Write2Excel $WorkSheet $Row 39 "URL"                       # AQ    
            $Col++
            Write2Excel $WorkSheet $Row $Col $Today "date"                # AR (Col 44)

            $ObjCount = ($Obj|Get-Member -Type NoteProperty).count
            statusmsg "$ObjCount data points contained in library." "red"            
            statusmsg "$col columns written" "Magenta"  

        }Else{
            StatusMsg "   --- No SNMPv3 ---" "Red"
            $Col = 10
            Write2Excel $WorkSheet $Row 12 "False"                      # J --[ Set FTP cell to "false" as well ]--
            Write2Excel $WorkSheet $Row 11 $Obj.SNMPv1                  # K
            $Worksheet.Cells($Row, 10).Font.ColorIndex = 0
            $Worksheet.Cells($Row, 10).Font.Bold = $False
            $Worksheet.Cells($Row, 12).Font.ColorIndex = 3              #--[ Apply red emphasis if SMNPv3 is not enabled ]--
            $Worksheet.Cells($Row, 12).Font.Bold = $true
            Write2Excel $WorkSheet $Row 12 $Obj.SNMPv3                  # L
        }
    }Else{ 
        Write2Excel $WorkSheet $Row 1 $Target 
        Write2Excel $WorkSheet $Row 2 $Obj.Hostname 
        StatusMsg "   --- No Connection ---" "Red"
        $Worksheet.Cells($Row, 9).Font.ColorIndex = 3
        $Worksheet.Cells($Row, 9).Font.Bold = $true
        $WorkSheet.Cells.Item($Row, 9) = "No Connection"
        #--[ Set row interior color as grey if device is OFFLINE prior to moving to next device ]-----
        $WorkSheet.UsedRange.Rows.Item($Row-1).Interior.ColorIndex = 15
    #---------------------------------------------------------------------------------------------
    }  

    #$RGB_Low = 226 + (239 * 256) + (219 * 256 * 256 )      #--[ Light Green ]--
    #$RGB_Low = 190 + (222 * 256) + (188 * 256 * 256)     #--[ Light Green ]--
    #$RGB_Medium = 245 + (219 * 256) + (151 * 256 * 256);
    #$RGB_High = 242 + (142 * 256) + (142 * 256 * 256);

    $Range = $WorkSheet.Range(("A$Row"),("AR$Row"))
    $Range.HorizontalAlignment = -4131
    1..4 | ForEach {
        $Range.Borders.Item($_).LineStyle = 1
        $Range.Borders.Item($_).Weight = 2
    }
    #--[ Set row interior color as green if device is online prior to moving to next device ]-----
    If ($Obj.Connection -eq "Online"){
        $WorkSheet.UsedRange.Rows.Item($Row-1).Interior.ColorIndex = 35
    }
    #---------------------------------------------------------------------------------------------
    $Resize = $WorkSheet.UsedRange
    [Void]$Resize.EntireColumn.AutoFit() 
    $Row++          
}

#--[ Cleanup ]--
Write-host ""
Try{ 
    If ((Test-Path -Path $ListFileName) -And (Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy")){
        StatusMsg 'Saving as "NewSpreadsheet.xlsx" ...' "Green"
        $Workbook.SaveAs("$PSScriptRoot\NewSpreadsheet.xlsx")
    }ElseIf(!(Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy")){
        StatusMsg "Saving as a new working spreadsheet... " "Green"
        $Workbook.SaveAs("$PSScriptRoot\$ExcelWorkingCopy")
    }Else{
        StatusMsg "Saving working spreadsheet... " "Green"
        $WorkBook.Close($true) #--[ Close workbook and save changes ]--
    }
    $Excel.quit() #--[ Quit Excel ]--
    [Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) #--[ Release the COM object ]--
}Catch{
    StatusMsg "Save Failed..." "Red"
}

Write-Host `n"--- COMPLETED ---" -ForegroundColor red

<#--[ XML File Example -- File should be named same as the script ]--
<!-- Settings & configuration file -->
<Settings>
    <General>
        <SourcePath>'C:\Users\Bob\Documents'</SourcePath>
        <ExcelSourceFile>Device-Master-Inventory.xlsx</ExcelSourceFile>
        <DNS>10.10.10.10</DNS>
    </General>
    <Credentials>
    <SMNPv3User>snmpuser</SMNPv3User>
        <SMNPv3AltUser>SNMPUser2</SMNPv3AltUser>
        <SNMPv3Secret>SnMpv3Community/SNMPv3Secret>
    </Credentials>
</Settings>    
#>

<#--[ APC OID Reference ]--
    "BattFreqOut" = ".1.3.6.1.4.1.318.1.1.1.4.2.2.0"
    "BattVOut" = ".1.3.6.1.4.1.318.1.1.1.4.2.1.0"
    "BattVIn" = ".1.3.6.1.4.1.318.1.1.1.3.2.1.0"
    "BattFreqIn" = ".1.3.6.1.4.1.318.1.1.1.3.2.4.0"
    "BattActualV" = ".1.3.6.1.4.1.318.1.1.1.2.2.8.0"
    "BattCurrentAmps" = ".1.3.6.1.4.1.318.1.1.1.2.2.9.0"
    "BattLastRepl" = ".1.3.6.1.4.1.318.1.1.1.2.1.3.0"
    "BattCapLeft" = ".1.3.6.1.4.1.318.1.1.1.2.2.1.0"
    "BattRunTime" = ".1.3.6.1.4.1.318.1.1.1.2.2.3.0"
    "BattReplace" = ".1.3.6.1.4.1.318.1.1.1.2.2.4.0"
    "BattTemp" = ".1.3.6.1.4.1.318.1.1.1.2.2.2.0"
    "ACVIn" = ".1.3.6.1.4.1.318.1.1.1.3.2.1.0"
    "ACFreqIn" = ".1.3.6.1.4.1.318.1.1.1.3.2.4.0"
    "LastXfer" = ".1.3.6.1.4.1.318.1.1.1.3.2.5.0"
    "UPSVOut" = ".1.3.6.1.4.1.318.1.1.1.4.2.1.0"
    "UPSFreqOut" = ".1.3.6.1.4.1.318.1.1.1.4.2.2.0"
    "UPSOutLoad" = ".1.3.6.1.4.1.318.1.1.1.4.2.3.0"
    "UPSOutAmps" = ".1.3.6.1.4.1.318.1.1.1.4.2.4.0"
    "LastTestResult" = ".1.3.6.1.4.1.318.1.1.1.7.2.3.0"
    "LastTestDate" = ".1.3.6.1.4.1.318.1.1.1.7.2.4.0"
    "BIOSSerNum" = ".1.3.6.1.4.1.318.1.1.1.1.2.3.0"
    "UPSModel" = ".1.3.6.1.4.1.318.1.1.1.1.1.1.0"
    "FirmwareVer" = ".1.3.6.1.4.1.318.1.1.1.1.2.1.0"
    "MfgDate" = ".1.3.6.1.4.1.318.1.1.1.1.2.2.0"
    "Location" = ".1.3.6.1.2.1.1.6.0"
    "Contact" = ".1.3.6.1.2.1.1.4.0"       
#>

