##########################################################
#
#  Version:        1.7 - 30-7-2024
#  Author:         MikeMaze84
#  Creation Date:  Tuesday, May 28, 2024 15:00 
#  Purpose/Change: Initial script development
#
##########################################################

# common var's
$logFile = $env:LOCALAPPDATA
$logfile = $logfile + "\temp\Company_Mapped_network_locations.log"
$logMessages = ""
$errTimeStamp = date

# Write Log File function
function write-log
    {
        param(
            [string]$logtext
            )
     ForEach ($logline in $logtext){
    add-content -Path $logFile $logline
    }
}

# Remove old logfile
$logMessages = @(
    "Intune - MS365 Drive Mappings Tool.",
    "Mike Maze",
    "-------------------------------------",
    "Execution time and date: $errTimeStamp",
    "Checking for old log"
    )

$testLogfile = Test-Path -Path $logFile
    if ( $testLogfile -ne $null ) {
        Remove-Item -Path $logFile -ErrorAction SilentlyContinue
        write-log -logtext $logMessages
        write-log -logtext "Old Log FIle removed."
       }

write-log -logtext $logMessages
write-log -logtext "Creating Function : Test-RunningAsSystem"
#check if running as system function
function Test-RunningAsSystem {
	[CmdletBinding()]
	param()
	process {
		return [bool]($(whoami -user) -match "S-1-5-18")
	}
}
write-log -logtext "Creating Function : Add-NetworkLocation"
# Add Network Location Function
function Add-NetworkLocation
{
    param(
        [string]$name,
        [string]$targetPath
    )
    
    # Get the basepath for network locations
    $shellApplication = New-Object -ComObject Shell.Application
    $nethoodPath = $shellApplication.Namespace(0x13).Self.Path

    # Only create if the local path doesn't already exist & remote path exists
    if ((Test-Path $nethoodPath) -and !(Test-Path "$nethoodPath\$name") -and (Test-Path $targetPath))
    {
        # Create the folder
        $newLinkFolder = New-Item -Name $name -Path $nethoodPath -type directory

        # Create the ini file
        $desktopIniContent = @"
[.ShellClassInfo]
CLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}
Flags=2
ConfirmFileOp=1
"@
        $desktopIniContent | Out-File -FilePath "$nethoodPath\$name\Desktop.ini"

        # Create the shortcut file
        $shortcut = (New-Object -ComObject WScript.Shell).Createshortcut("$nethoodPath\$name\target.lnk")
        $shortcut.TargetPath = $targetPath
        $shortcut.IconLocation = "%SystemRoot%\system32\SHELL32.DLL, 85"
        $shortcut.Description = $targetPath
        $shortcut.WorkingDirectory = $targetPath
        $shortcut.Save()
        
        # Set attributes on the files & folders
        Set-ItemProperty "$nethoodPath\$name\Desktop.ini" -Name Attributes -Value ([IO.FileAttributes]::System -bxor [IO.FileAttributes]::Hidden)
        Set-ItemProperty "$nethoodPath\$name" -Name Attributes -Value ([IO.FileAttributes]::ReadOnly)
    }
}
# Connection Test
write-log -logtext "Checking if the connection to the Company network is available."
write-log -logtext "Probing Server srvvbrnfs01."
$i = 0

for ($counterLoop = 0; $counterloop -lt 3; $counterLoop++) {
$i++
Start-Sleep -Seconds 5
$getConnectionStatus = Test-Connection -ComputerName "YOURFILESERVER.corp.domain.com" -ErrorAction SilentlyContinue
$errTimeStamp = date

if ($getConnectionStatus -eq $null){
   
    write-log -logtext "No Server Connection:$errTimeStamp"
    write-log -logtext "Stopping Script"
    # Wait state 60 Seconden  
    }
    else
    {
    write-log -logtext "Connection found : $getConnectionStatus"
    write-log -logtext "Removing old network shortcurs"
    $NewShellApplication = New-Object -ComObject Shell.Application
    $oldshortcuts = $NewshellApplication.Namespace(0x13).Self.Path
    write-log -logtext "Location : $oldshortcuts"
    $oldmappings = Get-ChildItem -Path $oldshortcuts
    foreach ($mapping in $oldmappings) {
        write-log -logtext "Set File Attributes to normal : $oldshortcuts\$mapping\Desktop.ini"
        Set-ItemProperty "$oldshortcuts\$mapping\Desktop.ini" -Name Attributes -Value ([IO.FileAttributes]::Archive -bxor [IO.FileAttributes]::Normal)
        write-log -logtext "Set File Attributes to normal : $oldshortcuts\$mapping"
        Set-ItemProperty "$oldshortcuts\$mapping" -Name Attributes -Value ([IO.FileAttributes]::Normal)
        write-log -logtext "removing : $mapping"
        Remove-Item -Path "$oldshortcuts\$mapping\*.lnk"
        Remove-Item -Path "$oldshortcuts\$mapping\Desktop.ini"
        Remove-Item -Path "$oldshortcuts\$mapping"
     }
## Drive D srvvbrnfs01
$user_id = whoami -upn
write-log -logtext "Checking for mappings for:"
write-log -logtext "$user_id"
write-log -logtext "\\YOURFILESERVER.corp.domain.com\data$"
$dirs = Get-childItem -Path "\\YOURFILESERVER.corp.domain.com\data$" -Directory -ErrorAction SilentlyContinue
foreach ($directory in $dirs) {
       $errTimeStamp = date
    write-log -logtext "Adding network location $directory at :$errTimeStamp"
    Add-NetworkLocation -name $directory -targetPath "\\YOURFILESERVER.corp.domain.com\$directory"
    write-log -logtext "Added"
    }
## Drive E srvvbrnfs01
write-log -logtext "\\YOURFILESERVER.corp.domain.com\data-e$"
$dirs = Get-childItem -Path "\\YOURFILESERVER.corp.domain.com\data-e$" -Directory -ErrorAction SilentlyContinue
foreach ($directory in $dirs) {
        $errTimeStamp = date
    write-log -logtext "Adding network location $directory at :$errTimeStamp"
    Add-NetworkLocation -name $directory -targetPath "\\YOURFILESERVER.corp.domain.com\$directory"
    write-log -logtext "Added"
    }


    ## Drive Users GDPR FS02
write-log -logtext "\\fs02.corp.domain.com\gdpr$"
$dirs = Get-childItem -Path "\\fs02.corp.domain.com\gdpr$" -Directory -ErrorAction SilentlyContinue
    $errTimeStamp = date
    $directory = "gdpr"
    Add-NetworkLocation -name $directory -targetPath "\\fs02.corp.domain.com\gdpr$"
    write-log -logtext "Added"

$counterLoop = 3
}
if ($counterLoop -eq 3 ) {
    write-log -logtext "Mapped after : $i"
    } else {
    write-log -logtext "Looping Connection test, try : $i"
    }
}
write-log -logtext "Completed."

#!SCHTASKCOMESHERE!#

###########################################################################################
# If this script is running under system (IME) scheduled task is created  (recurring)
###########################################################################################
if (Test-RunningAsSystem) {
    
	Start-Transcript -Path $(Join-Path -Path $env:TEMP -ChildPath "IntuneNetworkLocationsScheduledTask.log")
	Write-Output "Running as System --> creating scheduled task which will run on user logon"

	###########################################################################################
	# Get the current script path and content and save it to the client
	###########################################################################################

	$currentScript = Get-Content -Path $($PSCommandPath)

	$schtaskScript = $currentScript[(0) .. ($currentScript.IndexOf("#!SCHTASKCOMESHERE!#") - 1)]

	$scriptSavePath = $(Join-Path -Path $env:ProgramData -ChildPath "intune-Networklocations")

	if (-not (Test-Path $scriptSavePath)) {
        write-log -logtext "Creating new Directory $scriptSavePath"

		New-Item -ItemType Directory -Path $scriptSavePath -Force
	}
    write-log -logtext "Creating PowerShell Script add-NetworkLocations.ps1"
	$scriptSavePathName = "Add-NetworkLocations.ps1"

	$scriptPath = $(Join-Path -Path $scriptSavePath -ChildPath $scriptSavePathName)

	$schtaskScript | Out-File -FilePath $scriptPath -Force
     write-log -logtext "Script Path $scriptPath"
     write-log -logtext "Script Task $schtaskScript"
	###########################################################################################
	# Create dummy vbscript to hide PowerShell Window popping up at logon
	###########################################################################################

	$vbsDummyScript = "
	Dim shell,fso,file

	Set shell=CreateObject(`"WScript.Shell`")
	Set fso=CreateObject(`"Scripting.FileSystemObject`")

	strPath=WScript.Arguments.Item(0)

	If fso.FileExists(strPath) Then
		set file=fso.GetFile(strPath)
		strCMD=`"powershell -nologo -executionpolicy ByPass -command `" & Chr(34) & `"&{`" &_
		file.ShortPath & `"}`" & Chr(34)
		shell.Run strCMD,0
	 End If
	"

	$scriptSavePathName = "IntuneNetworkLocation-VBSHelper.vbs"

	$dummyScriptPath = $(Join-Path -Path $scriptSavePath -ChildPath $scriptSavePathName)

	$vbsDummyScript | Out-File -FilePath $dummyScriptPath -Force

	$wscriptPath = Join-Path $env:SystemRoot -ChildPath "System32\wscript.exe"

	###########################################################################################
	# Register a scheduled task to run for all users and execute the script on logon
	###########################################################################################


    $schtaskName = "IntuneNetworkLocationsMapping"
if (Get-ScheduledTask -TaskName $schtaskName -ErrorAction SilentlyContinue) {
    # If it exists, remove it
    Unregister-ScheduledTask -TaskName $schtaskName -Confirm:$false
    write-log -logtext "Unregister Scheduled Task"
}
	$schtaskDescription = "Add available network locations"

	$trigger = New-ScheduledTaskTrigger -AtLogOn
	#Execute task in users context
	$principal = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545" -Id "Author"
	#call the vbscript helper and pass the PosH script as argument
	$action = New-ScheduledTaskAction -Execute $wscriptPath -Argument "`"$dummyScriptPath`" `"$scriptPath`""
	$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

	$null = Register-ScheduledTask -TaskName $schtaskName -Trigger $trigger -Action $action  -Principal $principal -Settings $settings -Description $schtaskDescription -Force

	Start-ScheduledTask -TaskName $schtaskName
    write-log -logtext "Register Scheduled Task"
	Stop-Transcript
}
