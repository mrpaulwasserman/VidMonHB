<#
.SYNOPSIS

Video Monitor HandBrake Converter - Windows PowerShell Script  
  - Video monitoring script to automatically transcode video files
  
.DESCRIPTION

  This script will find specified types of video files and process them
  through the HandBrake Command Line Interface.  Parameters allow script
  to define functionality such as input folders, output folders, type of
  video files to search for, etc.
    - This script will handle up to 999 files at a time
    - Script will attempt to parse name to determine output folder structure
      Sample Output Folder Below:
      C:\Output\TV Show Name\Season 01\TV Show Name - S01E01 - Episode Name
  
  File metadata (specifically Title) will be removed if the taglib-sharp.dll
  module has been downloaded and placed in the same folder with this script.
  Note - This is not required because HandBrake will typically not use transfer
         metadata during the conversion process.

.LINK
  Use the following link to download the HandBrakeCLI.exe program
  https://handbrake.fr/downloads2.php

.LINK
  Use the following link to download the taglib-sharp.dll (not required)
  https://www.dllme.com/dll/files/taglib-sharp_dll.html


.PARAMETER 0
    vidTypes : Video type extensions to process
    Default  : mkv,avi

.PARAMETER 1
    in      : Input folder to monitor for video files
    Default : $ENV:UserProfile+"\Downloads"

.PARAMETER 2
    out     : Output folder to store converted video files
    Default : "D:\MediaOutput"

.PARAMETER 3
    outSameAsin : Override the out parameter and send output to 
                  the same folder where the input file originated from
    Default     : false

.PARAMETER 4
    delAfterConv : Delete the original file that was convereted (normally .mkv) (Maintain, Delete or Recycle)
    Default      : Maintain

.PARAMETER 5
    $hbloc  : Directory location of HandBrake Command Line Interface program 
    Default :  $ENV:ProgramFiles+"\HandBrake\HandBrakeCLI.exe"

.PARAMETER 6
    $hbpreloc : Directory location of HandBrake presets 
    Default   : $ENV:APPDATA+"\HandBrake\presets.json"
    
.PARAMETER 7
    $tvPreset : HandBrake preset name to use (uses default preset from $hbpreloc file if none specified)
                NOTE - Preset name should have no spaces (sample preset name = VeryFastDDtoAAC)
    Default   : ""
    
.PARAMETER 8
    $moviePreset : HandBrake preset name to use (uses default preset from $hbpreloc file if none specified)
                   NOTE - Preset name should have no spaces (sample preset name = VeryFastDDtoAAC)
    Default      : ""

.PARAMETER 9
    $hbopts : Additional HandBrake options
    Default : ""

.PARAMETER 10
    $propfileloc : Directory location of properties file (if exists, it will replace parameters)
    Default      : "C:\VidMonHB\VidMonHB.ps-properties"

.PARAMETER 11
    $movefiles : Move converted file to final folder location
    Default    : false

.PARAMETER 12
    $logfilePath : Folder to store log files
    Default      : $out+"\logs"

.PARAMETER 13
    $TVShowBasePath : TV Show Output Folder (used for moving files)
    Default         : "D:\\Media\\02. TV Shows\\"

.PARAMETER 14 
    $movieBasePath : Movie Output Folder (used for moving files)
	  Default        : "D:\\Media\\"

.PARAMETER 15
    $ParallelProcMax : Parallel processing (convert multiple video files simultaneously)
    Default          : 0 or 1 indicates single threaded processing (Max 10)

.PARAMETER 16
    $limit  : Process the first x number of files found
    Default : 999

.PARAMETER 17
    $postExecCmd  : # Post command line execution. Can execute a specific command or a batch script

.PARAMETER 18
    $postExecArgs : # Post command line execution arguments

.PARAMETER 19
    $postNotify : All = Send All notifications
                : Error = Send Errors only
                : None = No notifications
    Default     : None

.PARAMETER 20
    $smtpServer : SMTP Relay Server

.PARAMETER 21
    $smtpFromEmail : From email address used by the SMTP Relay Server

.PARAMETER 22
    $smtpToEmail : To email address(es) used during post notification process

.PARAMETER 23
    $postLog : Always = Always open log file
               Error = Only open log if an error occurs
               Never = Never open the log file
    Default  : Always

.PARAMETER 24
    $repeatCtr : # of times to repeat this script
    Default  : 0

.PARAMETER 25
    $repeatMonitor : Continually monitor and repeat running when a file is found
    Default  : false

    .INPUTS

  VidMonHB.ps-properties


.OUTPUTS

  Logs:
  There are 2 types of log output that will be generated when this script runs.
    1. A summary log file (VidMonHB_YYYY-MM-DD_HH-MM-SS_Summary.txt)
      - The summary log file only contains output from this script
    2. A detail Handbrake log file (Filename_YYYY-MM-DD_HH-MM-SS_HBDetails.txt)
      - The detailed log file contains output from the handbrake CLI program

  A log cleanup is built into this script and will automatically delete any log
  files that are more than x days old (default is 30 days)

  Video Output:
    1. Sample TV Show Output Folder Below:
      - C:\Output\TV Show Name\Season 01\TV Show Name - S01E01 - Episode Name
  
.NOTES

  Author:  Paul Wasserman
  Purpose: PowerShell Video Monitor/Converter Script
  Author   Version Date        Description
  Paul W.   1.0    03/27/2020  Initial Version.
            1.2    04/01/2020  Properties file now working.
            1.3    04/02/2020  Add parallel processing capabilities.
            1.4    04/03/2020  Allow output to simply be created in the same folder 
                               as the originating input file.
                               NOTE - This setting overrides the $out paramemter.
            1.5    04/03/2020  Add check to ensure HandBrakeCLI program is installed.
            1.6    04/03/2020  Updated MoveFile to allow for different folder pattern.
            1.7    04/04/2020  Add in disk space usage information.
                               Add $limit parameter to allow user to specify # of files to process.
                               Add $postExecCmd & $postExecArgs parameter to allow post execution 
                               of a command line item
                               Add x of x processing info.
            1.8    04/04/2020  Fixed time calculation.
            1.9    04/05/2020  Moved parallel processing check immediately after job submission.
            1.10   04/05/2020  Added interactive logic to allow parameters to be entered.
                               NOTE - User entry will default timeout after 5 seconds.
                               Fixed outputlog location to use $logfilepath.
                               Added logic to scan HB logfile and prevent file deletion if job did 
                               not successfully complete.
                               Added alternate background colors for Error messages (White on 
                               Red background).
            1.11   04/06/2020  Add notification logic.
            1.12   04/07/2020  Change log extension from .log to .txt so emailed notifications can 
                               be opened on devices.
            1.13   04/11/2020  Major restructure to check job completion during processing. Script
                               will now remove original files as soon as possible, and not all at 
                               the very end. This will allow for better disk space management and 
                               future restart functionality.
            1.14   04/14/2020  Minor corrections and begin adding resume logic.
                   04/19/2020  Completed resume logic.
                               Add save config option.
                               Add time estimation based on total size being processed.
                               Add time diff after each completion.
                               Clear ReadOnly attribute of files that need to be deleted after 
                               conversion (with error notifications)
            1.15   04/20/2020  Minor corrections. Log spacing.
            1.16   04/26/2020  Major change... Switch to GUI based parameter screen!!
                               Separate presets into 2 separate options (TVPreset & MoviePreset).
                               Add SMTP Information.
                               Add Tooltips, but currently disabled because they working perfectly yet.
                               Add description information to each input field. 
                               Add send to Recycle Bin logic.
                                 Must set variable $recycle to $true.
                                 Must install Powershell Recycle module (see instructions below).
                               Renamed resume.ps-properties to resumeVidMonHB.txt.
                               Move COMPLETED message to bottom of block when not parallel processing.
                               Add ability to select from multiple config/parm files.
                                 Also add ability to write out config to a new file.
            1.17   05/02/2020  Added dropdown list for original file action.
                                 Delete, Recycle and Maintain.
                               Added dropdown list for post open log action.
                                 Always, Error, Never 
                              Change name of properties to .ps-properties (for Mark)
            1.18   05/03/2020 Add pre-execution verification logic
                              Add advanced movie folder logic (store movies by year)
            1.19   05/10/2020 Add additional verification logic to light up invalid fields
                              Corrections to pre-exectuion logic
            1.20   05/15/2020 Corrected delFile fname message
                              Corrected log file cleanup and changed to Recycle log files
            1.21   05/16/2020 PS 7.0.1 fix.  Removed using assembly microsoft.visualbasic
            1.22   05/25/2020 assembly microsoft.visualbasic required for Recycle to work.
                              Will need to figure something else out for 7.0.1.
                              Corrected $errorCount issue.
            1.23   06/01/2020 Small correction to correct slash.
            1.24   06/13/2020 Add logic to repeat n # of times and made some minor corrections.
            1.25   06/28/2020 Add colors to highlight disk savings or loss.
                              Requires PSWriteColor module to be installed
            1.26   07/02/2020 Minor string correction to ETA.
                              Add $padSize and correct padding msg (i.e. 0001 of 1000).
                              Rearrange start msg info a little bit.  Now processing on it's own line.
                              Add starting size to start time message.
                              Commented out Clearing metadata msg.
                              Added background highlighting around ***START and ****COMPLETED msgs.
            1.27   07/17/2020 Continually monitor folder and execute when files are found.
                              .\VidMonHB.ps1 -repeatmonitor $true
            1.28   07/22/2020 Add running history information $HistoryLogFile.
                              Display history summary during monitor mode
            1.29   08/30/2024 Provide option to skip smaller files.



  First time execution may require running the following command (for PowerShell 5 & lower)
    Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force

  If you want to recycle files (instead of delete), install this Powershell module
    Install-Module -Name Recycle -RequiredVersion 1.0.2 -Scope CurrentUser -Force

  Run the following install script to include the write command with colors.
    Install-Module -Name PSWriteColor -Scope CurrentUser -Force

  There are tooltips for each of the input fields. If these are not showing up, check the following
  Computer\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ 
  there needs to be a DWORD entry for EnableBalloonTips with the value set to 1

When the Move File option is set, converted files are moved to the approriate folders (see examples below)
  - Movie files are placed into date folders
  - TV Show files are placed into show name folders
  
\Media\00. Movies  is the base default location:

    \Media
    |-- 00. Movies 
    |   |-- 01. Movies 1920-1979
    |   |-- 02. Movies 1981-1999
    |   |-- 03. Movies 2000-2019
    |   |-- 04. Movies 2020-2029


\Media\02. TV Shows  is the base default location:

    \Media\02. TV Shows
    |-- 02. TV Shows
    |   |-- Dynasty
    |   |-- Green Acres
    |   |-- The Andy Griffith Show
    |   |-- The Flintstones


.EXAMPLE
  VidMonHB.ps1 -vidTypes avi -in "c:\mediaIn" -out "c:\mediaOut" -delAfterConv Delete
  --> Converts .avi files and delete the originals (using the specified in/out folder locations)
  
.EXAMPLE
  VidMonHB.ps1 -vidTypes avi -delAfterConv Recycle -hbopt -O
  --> Converts .avi files, Recycle the originals, and add the Optimize HandBrake option

.EXAMPLE
  VidMonHB.ps1 -vidTypes mkv -delAfterConv Maintain -movefiles True $TVShowBasePath "D:\\Media\\02. TV Shows\\""
  --> Converts .mkv files, maintain the originals, move the converted files to new base location

.EXAMPLE
  VidMonHB.ps1 -delAfterConv Maintain -tvPreset VeryFastDDtoAAC
  --> Converts mkv and avi files, uses a HandBrake preset named VeryFastDDtoAAC

.EXAMPLE
  VidMonHB.ps1 -repeatCtr 5
  --> Runs the VidMonHB script 5 times
#>

# Used for Recycle Bin logic
using assembly microsoft.visualbasic
using namespace microsoft.visualbasic

#Parameters - Make sure all parameters have a trailing comma except for the final one
Param
  (
    #[CmdletBinding()]
    # Type of video file(s) to process (i.e. mkv, avi) 
    [Parameter(Position=0)]
    [string]$vidTypes = "mkv,avi",

    # Input folder to monitor for video files
	  # This is typically changed during initial setup.
    [Parameter(Position=1)]
    [string]$in = $ENV:UserProfile+"\Downloads",

    # Output folder to store converted video files
	  # This is typically changed during initial setup.
    [Parameter(Position=2)]
    [string]$out = "D:\MediaOutput",

    # Override the out parameter and send output to 
    # the same folder where the input file originated from
    [Parameter(Position=3)]
    [Boolean]$outSameAsIn = $false,

    # Delete the original file that was convereted
    [Parameter(Position=4)]
    [ValidateSet("Maintain","Delete","Recycle")]
    $delAfterConv = "Maintain",

    # Directory location of HandBrake Command Line Interface program 
    [Parameter(Position=5)]
    [string]$hbloc = $ENV:ProgramFiles+"\HandBrake\HandBrakeCLI.exe",

    # Directory location of HandBrake presets 
    [Parameter(Position=6)]
    [string]$hbpreloc = $ENV:APPDATA+"\HandBrake\presets.json",
    
    # HandBrake preset name to use (uses default preset from $hbpreloc file if none specified)
    # NOTE - Preset name should have no spaces (sample preset name = VeryFastDDtoAAC)
    [Parameter(Position=7)]
    [string]$tvPreset = "VeryFastDDtoAAC",

    # HandBrake preset name to use (uses default preset from $hbpreloc file if none specified)
    # NOTE - Preset name should have no spaces (sample preset name = VeryFastDDtoAAC)
    [Parameter(Position=8)]
    [string]$moviePreset = "VeryFastDDtoAAC",

    # Additional HandBrake options
    [Parameter(Position=9)]
    [string]$hbopts = "",

    # Directory location of properties file (if exists, it will replace parameters)
    [Parameter(Position=10)]
    [string]$propfileloc = (".\VidMonHB.ps-properties"),
    
    # Move converted file to final folder location
    [Parameter(Position=11)]
    [Boolean]$movefiles = $false,

    # Folder to store log files
    [Parameter(Position=12)]
    [string]$logfilePath = $out+"\logs",

    # TV Show Output Folder (used for moving files)
	  # This is typically changed during initial setup.
    [Parameter(Position=13)]
    [string]$TVShowBasePath = "D:\Media\02. TV Shows\",

    # Movie Output Folder (used for moving files)
	  # This is typically changed during initial setup.
    [Parameter(Position=14)]
    [string]$movieBasePath = "D:\Media\",

    # Parallel processing (convert multiple video files simultaneously)
    # Note - The higher the #, the more Memory & CPU will be used. Be careful.
    # Default 0 indicates single threaded processing.
    [Parameter(Position=15)]
    [int]$ParallelProcMax = 0,

    # Process the first x number of files found
    # Default 999
    [Parameter(Position=16)]
    [int]$limit = 999,

    # Post command line execution.  Can execute a specific command or a batch script
    [Parameter(Position=17)]
    [string]$postExecCmd = "",

    # Post command line execution arguments
    [Parameter(Position=18)]
    [string]$postExecArgs = "",

    # Post command line execution arguments
    # All=Send All notifications
    # Error=Send Errors only
    # None=Do not send notifications
    [Parameter(Position=19)]
    [ValidateSet("None","All","Error")]
    [string]$postNotify = "",

    # SMTP Relay Server
    [Parameter(Position=20)]
    [string]$smtpServer = "",

    # SMTP From email address used by the SMTP Relay Server
    [Parameter(Position=21)]
    [string]$smtpFromEmail = "",

    # SMTP To email address(es) used during post notification process
    [Parameter(Position=22)]
    [string]$smtpToEmail = "",

    # Post log - Should the log file be opened when the script completes
    # Always = Always open log file
    # Error = Only open log if an error occurs
    # Never = Never open the log file
    [Parameter(Position=23)]
    [string]$postLog = "Always",

    # of times to repeat this script
    [Parameter(Position=24)]
    [int]$repeatCtr = 0,

    # Continually monitor and repeat running when a file is found
    [Parameter(Position=25)]
    [Boolean]$repeatMonitor = $false,

    # Look for files larger than specified size
    [Parameter(Position=26)]
    [string]$minSize = "0gb"

  )

#-------------------------------------------[Declarations]-------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#Script Version
$version="1.29"
$beginTime=Get-Date

#Parm/Config entry type (Windows form or Powershell entry)
$entryType="winForm"

$videoFiles=@()
$procObject=@() # Array of id, begTime, endTime, begSize, endSize, detailLogFile,
#                            baseName, fullName, newFileName, processed
$procList=@()   # Array of $procObject

#Used for log file naming
$timestamp=(Get-Date -Format "yyyy-MM-dd_HH-mm-ss")

#Set Error Action to Silently Continue
$ErrorActionPreference="SilentlyContinue"
$errorCount=0

$totBegSize=0; $totEndSize=0; $totSizDiff=0

$currentEnv=Convert-Path "."
$vidMonHBImage = $currentEnv + "\imageSmall.png"

#Store the parameters in a hash table in case of a resume
$resumeParms=@{}
$resumeFile=($currentEnv + "\resumeVidMonHB.txt")
$resume=$false

$propfileloc = Convert-Path $propfileloc

$readonly=[System.IO.FileAttributes]::ReadOnly
$readOnlyErrCnt=@()

[boolean]$exit = $true
#SMTP Information
$serverName = "PlexDad"

[int]$padSize = 3  #Default to 3

#Form variables
$blue = [System.Drawing.Color]::FromArgb(0,120,250)
$cyan = [System.Drawing.Color]::FromArgb(0,255,255)
$lightblue = [System.Drawing.Color]::FromArgb(173,216,230)
$white = [System.Drawing.Color]::FromArgb(255,255,255)
$red = [System.Drawing.Color]::FromArgb(255,0,0)
$yellow = [System.Drawing.Color]::FromArgb(255,255,0)
$lblColor = $yellow
$logBGcolor=$null

$v9bi = New-Object System.Windows.Forms.Label
$v9bi.font = (new-object System.Drawing.Font('Verdana',9,[System.Drawing.FontStyle]::Bold))
$v12 = New-Object System.Windows.Forms.Label
$v12.font = (new-object System.Drawing.Font('Verdana',12,[System.Drawing.FontStyle]::Regular))
$v12b = New-Object System.Windows.Forms.Label
$v12b.font = (new-object System.Drawing.Font('Verdana',12,[System.Drawing.FontStyle]::Bold))
$v12bi = New-Object System.Windows.Forms.Label
$v12bi.font = (new-object System.Drawing.Font('Verdana',12,[System.Drawing.FontStyle]::Bold))
$v14b = New-Object System.Windows.Forms.Label
$v14b.font = (new-object System.Drawing.Font('Verdana',14,[System.Drawing.FontStyle]::Bold))
$v16bi = New-Object System.Windows.Forms.Label
$v16bi.font = (new-object System.Drawing.Font('Verdana',16,[System.Drawing.FontStyle]::Bold))
$v30b = New-Object System.Windows.Forms.Label
$v30b.font = (new-object System.Drawing.Font('Verdana',30,[System.Drawing.FontStyle]::Bold))
$ss12 = New-Object System.Windows.Forms.Label
$ss12.font = (new-object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]::Regular))

$bs3 = [System.Windows.Forms.BorderStyle]'Fixed3D'

#--------------------------------------------[Functions]---------------------------------------------

#Function to clear the Title metadata
function clearTitleMeta {
  Param ([string[]]$fileName)
  if ( -not ($ClearMetaFlag)) {return}
  #writeLog "Clearing metadata from file: $filename" -logOnly "L" 
  $fileName = Get-ChildItem -Path $fileName
  $mediaFile = [TagLib.File]::Create($fileName.fullName)
  [TagLib.Mpeg4.AppleTag]$customTag = $mediaFile.GetTag([TagLib.TagTypes]::Apple, 1)
  $customTag.Title = ""
  $mediaFile.Save()
}

#Function to write out log information (to logFile and screen)
function writeLog ($logMsg,$logType,$logSeverity,$logBGcolor)
{
  if ($Host.Name -eq "Visual Studio Code Host") {$bgColor="Black"}
  else {$bgColor = [System.Console]::BackgroundColor}
  $fgColor = "Yellow"
  if ($null -ne $logBGcolor) {$bgColor = $logBGcolor; $logBGcolor=$null} #Override
  if ($logSeverity -eq "E") {$fgColor="White";$bgColor="Red"}
  switch ($logType) {
    "L" { Write-Output $logMsg | Out-File $sumLogFile -Append }
    "S" { Write-Host $logMsg -ForegroundColor $fgColor -BackgroundColor $bgColor }
    Default {
      Write-Output $logMsg | Out-File $sumLogFile -Append
      Write-Host $logMsg -ForegroundColor $fgColor -BackgroundColor $bgColor
    }
  }
}

#Try to determine if this is a TV show or a Movie
function checkIfTVfile ($fileName) {
  $season = ""; $episode = ""; $epiName = ""; $folder = ""
  $season =  ($fileName -split ('S*(\d{1,2})(x|E)(\d{1,2})'))[1]
  $episode = ($fileName -split ('S*(\d{1,2})(x|E)(\d{1,2})'))[3]
  $epiName = ($fileName -replace '.*\\')
  $folder =  ($epiName  -split ('S*(\d{1,2})(x|E)(\d{1,2})'))[0]; 
  $folder = $folder.Substring(0, $folder.Length - 1).Trim()
  $folder = $folder.TrimEnd("-").Trim()
  if ($folder.ToUpper() -contains "TV" -or ($null -ne $season+$episode))
  { return $true }
  else { return $false }
}

#Determine if this is a Movie file or a TV file, then call the appropriate move function
function moveFile ($fileName)
{
  if ( -not ($movefiles)) {return}
    if (checkIfTVfile($filename)) { 
      moveTVFile($fileName) }
    else { moveMovieFile($fileName) }
}

#Function to move Movie files to new location
function moveMovieFile ($fileName)
{
  #Look for a Movies folder by year to move file to.
  $movieYear = ($fileName -split ('\(([^\)]+)\)'))[1]
  $movieFolders = get-childitem $movieBasePath -recurse -Directory -Include "*movie*" | Sort-Object
  #$defaultdestPath = $movieFolders.FullName[0] + "\"  
  if ( ($movieFolders | Measure-Object).Count -eq 1) { 
    $destPath = $movieFolders.FullName + "\" + (Split-Path $filename -leaf)
    writeLog ("Movie - Move-Item -Path $fileName -Destination $movieBasePath -Force")
    Move-Item -Path $fileName -Destination $movieBasePath -Force 
    return
  }
  if ($null -ne $movieYear) {
    foreach ($movieFolder in $movieFolders) {
      $moviePathMinYear = ($movieFolder.Name -split ('(\d{4})( ?- ?)?(\d{4})?'))[1]
      $moviePathMaxYear = ($movieFolder.Name -split ('(\d{4})( ?- ?)?(\d{4})?'))[3]
      if (($movieFolder.Name -contains $movieYear) -or 
          ($movieYear -eq $moviePathMinYear) -or
          ($movieYear -in $moviePathMinYear .. $moviePathMaxYear)
          ) 
      {
        $destPath = $movieFolder.FullName + "\" + (Split-Path $filename -leaf)
        writeLog ("Movie - Move-Item -Path $fileName -Destination " + $destPath + " -Force")
        Move-Item -Path $fileName -Destination $destPath -Force
        return
      }
    }
  }
  $destPath = $movieBasePath + (Split-Path $filename -leaf)
  writeLog ("Movie - Move-Item -Path $fileName -Destination $destPath -Force")
  Move-Item -Path $fileName -Destination $destPath -Force
}

#Function to move TV Show files to new location
function moveTVFile ($fileName) {
  #Parse out the folder name
  $season = ""; $episode = ""; $epiName = ""; $folder = ""; $newPathBase = ""; $newPath = "";
  $season =  ($fileName -split ('S*(\d{1,2})(x|E)(\d{1,2})'))[1]
  $episode = ($fileName -split ('S*(\d{1,2})(x|E)(\d{1,2})'))[3]
  $epiName = ($fileName -replace '.*\\')
  $folder =  ($epiName  -split ('S*(\d{1,2})(x|E)(\d{1,2})'))[0]
  $folder = $folder.Substring(0, $folder.Length - 1).Trim()
  $folder = $folder.TrimEnd("-").Trim()
  if ($season -eq "" -or $episode -eq "") {
    writeLog ("Cannot move file because filename is missing season or episode info")
    return
  }
  $newPathBase = ($TVShowbasePath + $folder + "\" + "Season " + $season)
  $newPath = ($newPathBase + "\" + $epiName)
  #If the destination directory doesn't exist, create it
  if ( -not (Test-Path $newPathBase)) { New-Item $newPathBase -ItemType Directory -Force }
  writeLog ("TV Show - Move-Item -Path $fileName -Destination $newPath -Force") #-logFlag "L"
  Move-Item -Path $fileName -Destination $newPath -Force
}

#Automatically clean up log files that are greater than x number of days old (Extra check logic for safety).
function cleanoldLogs ([int]$numDays) {
  if ($null -eq $numDays) { $numDays = 30 }
  $oldLogs = Get-ChildItem $logFilePath | Where-Object {$_.extension -in ".log",".txt"} | Where-Object CreationTime -lt (Get-Date).AddDays(-$numDays)
  $num = 0; $num = ($oldLogs | Measure-Object).Count
  if ($null -eq $oldLogs -or $num -eq 0 -or $null -eq $logFilePath) { 
    writeLog ("`nNo logs to clean up"); return 
  }
  else {
    writeLog ("`nNow cleaning up $num old log file(s)")
  }    
  foreach ($oldLog in $oldLogs) { 
    if ($recycleAvailble) { recycleFile($oldLog.FullName) } 
    else { delFile($oldLog.FullName) } 
  }
  writeLog ("Cleaned up $num log file(s) that were > $numDays old")
}

#Check if a file is open/in use
function checkFileLocked($filePath) {
  $fileInfo = New-Object System.IO.FileInfo $filePath
  try {
    $fileStream = $fileInfo.Open( [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read )
    return $false }
  catch { return $true } #file is in use 
}

function displayParms($logType) {
  writeLog "`n`n`t`t`tVidMonHB - Version $version" -logType $logType
  writeLog (("-").PadRight(80,"-")) -logType $logType
  writeLog ("Parameter Settings") -logType $logType
  writeLog (" 1) Video types             : $vidTypes") -logType $logType
  writeLog (" 2) Input location          : $in") -logType $logType
  writeLog (" 3) Output location         : $out" ) -logType $logType
  writeLog (" 4) Same As Input Override  : $outSameAsIn") -logType $logType
  writeLog (" 5) Delete original         : $delAfterConv") -logType $logType
  writeLog (" 6) HandBrake file location : $hbloc") -logType $logType
  writeLog (" 7) HB preset file location : $hbpreloc") -logType $logType
  writeLog (" 8) HandBrake TV preset     : $tvPreset") -logType $logType
  writeLog (" 9) HandBrake Movie preset  : $moviePreset") -logType $logType
  writeLog ("10) HandBrake options       : $hbopts") -logType $logType
  writeLog ("11) Property location       : $propfileloc") -logType $logType
  writeLog ("12) Move files option       : $movefiles") -logType $logType
  writeLog ("13) Log file path           : $logfilePath") -logType $logType
  writeLog ("14) TV Show base path       : $TVShowBasePath") -logType $logType
  writeLog ("15) HandBrake Movie preset  : $moviePreset") -logType $logType
  writeLog ("16) Movie base path         : $movieBasePath") -logType $logType
  switch ($ParallelProcMax) { {$_ -lt 2} {$ppmsg="Single Threaded Mode"  } Default {$ppmsg="Parallel Processing Mode Enabled"} }
  writeLog ("17) Parallel processing max : $ParallelProcMax - $ppmsg") -logType $logType
  writeLog ("18) # of files to process   : $limit") -logType $logType
  writeLog ("19) Post exec cmd           : $postExecCmd") -logType $logType
  writeLog ("20) Post exec arguments     : $postExecArgs") -logType $logType
  writeLog ("21) Post notify (All,Error) : $postNotify") -logType $logType
  writeLog ("22) SMTP Relay Server       : $smtpServer") -logType $logType
  writeLog ("23) Email From              : $smtpFromEmail") -logType $logType
  writeLog ("24) Email To                : $smtpToEmail") -logType $logType
  writeLog ("") -logType $logType
  writeLog ("77) Save config") -logType $logType
  writeLog ("88) Exit script") -logType $logType
  writeLog ("99) Execute script (default)") -logType $logType
  writeLog ("") -logType $logType
}


#Pass in (by ref) table procList. Finalize any completed jobs and update the processed flag
function chkForCompletion($jobList) {
  if ($ParallelProcMax -gt 1 -and $fileCount -gt 1) {
    if (($jobList.processed | where-object {$_ -eq $false} | Measure-Object).Count -gt 0) {
      writeLog ("`nIn Progress") -logType "S"
      foreach ($job in $joblist) {
        if ($job.processed -eq $false) {
          writeLog ($job.countMsg + " - " + $job.baseName) -logType "S"
        }
      }
    }
    writeLog ("") -logType "S"
  }

  foreach ($job in $jobList) { 
    $chkJob = Get-Process -id $job.id -ErrorAction SilentlyContinue
    #If a Job has completed, perform final processing logic
    if ($job.processed -eq $false -and $chkJob.HasExited -in ($null,$true)) { 
      $job.processed = $true
      $job.endTime = Get-Date
      $job.endSize = [math]::Round((get-item $job.newFileName).Length / 1GB,3)
      if ($ParallelProcMax -gt 1 -and $fileCount -gt 1) {
        writeLog "****COMPLETED****COMPLETED****COMPLETED****COMPLETED****COMPLETED****COMPLETED****COMPLETED****COMPLETED****" -logBGcolor $msgBGcolor
      }
      writeLog ("Completed : " + $job.countMsg + " - """ + $job.newFileName + """")
      writeLog ("Log file  : " + $job.dtlLogFile)
      $timeDiff = getTimeDiff $job.begTime $job.endTime
      # writeLog ("Start time: " + $job.begTime.ToString() + 
      #          "   End time: " + $job.endTime.ToString() +
      #        "   Total time: " + $timeDiff.Hours + " hrs  " + 
      #          $timeDiff.Minutes + " mins  " + $timeDiff.Seconds + " secs")
      if ($timeDiff.Hours -eq 0) {
        $msgTime = "   Total time: " + $timeDiff.Minutes + " mins  " + $timeDiff.Seconds + " secs" }
      else {
        $msgTime = "   Total time: " + $timeDiff.Hours + " hrs  " + 
                   $timeDiff.Minutes + " mins  " + $timeDiff.Seconds + " secs"
      }        
      writeLog ("End time  : " + $job.endTime.ToString() + $msgTime)
      # $sizeInfo = "Start size: " + [math]::Round($job.begSize,3) + " GB   End size: " +
      #             [math]::Round($job.endSize,3) + " GB   "
      $sizeInfo = "End size  : " + [math]::Round($job.endSize,3) + " GB  "
      $diskSavings = [math]::Round(($job.begSize - $job.endSize),3) 
      $diskSavingsPCT = [string]([math]::Round(100-($job.endSize / $job.begSize)*100,2)) + "%"
      writeLog ($sizeInfo + "Disk savings: " + $diskSavings + " GB  " + $diskSavingsPCT) -logType "L"
      if ($diskSavings -ge 0) {
        Write-Color $sizeInfo, "Disk savings: ", $diskSavings, " GB  ", $diskSavingsPCT -Color White, Black, Black, Black, Black -BackGroundColor $bgColor, Green, Green, Green, Green }
      else {        
        Write-Color $sizeInfo, "Disk loss: ", $diskSavings, " GB" -Color White, White, White, White -BackGroundColor $bgColor, Red, Red, Red }
      # Now set the file in the resume file as complete
      $replacethis = "Unprocessed videofile="+$job.fullName
      $replacethat = "Completed videofile="+$job.fullName
      (Get-Content $resumeFile).replace($replacethis,$replacethat) | Set-Content $resumeFile
                
      if($delAfterConv -eq "Maintain") {writeLog ("Maintained original file " + $job.fullName)}
      else {
        #Check HB log file to ensure completion before removing original file (Must have 'Finished work at')
        #Search $proc for log file name
        #$findLogFile = $procList | Where-Object {$_.baseName -eq $file.baseName} 
        $count = @( Get-Content $job.dtlLogFile | Where-Object { $_.Contains("Finished work at") } ).Count
        if ($count -eq 1) {
          #First check if attribute of the file we're tring to delete is ReadOnly. If yes, try to clear it
          $chkFile=Get-ChildItem $job.fullName
          if (($chkFile).Attributes -band $readonly -eq $readonly) {
            #attrib -r $chkFile
            ($chkFile).Attributes -= $readonly.value__
            if (($chkFile).Attributes -band $readonly -eq $readonly) {
              writeLog ("Error - Can't delete file because ReadOnly Attribute could not be cleared for " + $job.fullName) -logSeverity "E"
              $readOnlyErrCnt += $job.fullName
            }
          }
          delFile($job.fullName)
        }
        else { writelog ("HandBrake error found. Original file was not removed.`nPlease review log " + 
                        $job.dtlLogFile) -logSeverity "E" 
               $script:errorCount += 1 
        }
      } #$delAfterConv
      clearTitleMeta($job.newFileName)
      moveFile ($job.newFileName)
      if ($ParallelProcMax -lt 2) {
        writeLog "****COMPLETED****COMPLETED****COMPLETED****COMPLETED****COMPLETED****COMPLETED****COMPLETED****COMPLETED****" -logBGcolor $msgBGcolor
      }
      writeLog "" 
    } #if 
  } #foreach
}

#Return the difference between two times
function getTimeDiff($beginTime,$endTime) {
  $timeDiff = New-TimeSpan $beginTime $endTime
  if ($TimeDiff.Seconds -lt 0) {
    $timeDiff = New-TimeSpan $endTime $beginTime
  }
    return $timeDiff
}

#Read through the original configuration file and update with the new values
function saveConfig() {

  # Check if the file exists. 
  if (test-path $propfileloc) {$newParms = (Get-Content $propfileloc) }
  else {
    try {$newParms = (Get-Content ".\VidMonHB.ps-properties")}
    catch {
      $txt_Info.AppendText("Error - Original VidMonHB.ps-properties file is required to create a new config file.")
      $txt_Info.AppendText("`n")
      #Scroll to the end of the textbox
      $txt_Info.SelectionStart = $txt_Info.TextLength;
      $txt_Info.ScrollToCaret()
      return
    }
  }
  $i = 0
  foreach ($object in $newParms) {
    $check = $object.Split("=")[0]
    switch (($check).ToUpper()) {
      "vidTypes" { $newParms[$i] = "vidTypes=$vidTypes" }
      "IN" { $newParms[$i] = "in=" + ($in.Replace("\","\\")) }
      "OUT" { $newParms[$i] = "out=" + ($out.Replace("\","\\")) }
      "OUTSAMEASIN" { $newParms[$i] = "outSameAsIn=$outSameAsIn" }
      "DELAFTERCONV" { $newParms[$i] = "delAfterConv=$delAfterConv" }
      "HBLOC" { $newParms[$i] = "hbloc=" + ($hbloc.Replace("\","\\")) }
      "HBPRELOC" { $newParms[$i] = "hbpreloc=" + ($hbpreloc.Replace("\","\\")) }
      "TVPRESET" { $newParms[$i] = "tvPreset=$tvPreset" }
      "MOVIEPRESET" { $newParms[$i] = "moviePreset=$moviePreset" }
      "HBOPTS" { $newParms[$i] = "hbopts=$hbopts" }
      "MOVEFILES" { $newParms[$i] = "movefiles=$movefiles" } 
      "LOGFILEPATH" { $newParms[$i] = "logfilePath=" + ($logfilePath.Replace("\","\\")) }
      "TVSHOWBASEPATH" { $newParms[$i] = "TVShowBasePath=" + ($TVShowBasePath.Replace("\","\\")) }
      "MOVIEBASEPATH" { $newParms[$i] = "movieBasePath=" + ($movieBasePath.Replace("\","\\")) }
      "PARALLELPROCMAX" { $newParms[$i] = "ParallelProcMax=$ParallelProcMax" }
      "LIMIT" { $newParms[$i] = "limit=$limit" }
      "POSTEXECCMD" { $newParms[$i] = "postExecCmd=" + ($postExecCmd.Replace("\","\\")) }
      "POSTEXECARGS" { $newParms[$i] = "postExecArgs=$postExecArgs" }
      "POSTNOTIFY" { $newParms[$i] = "postNotify=$postNotify" }
      "SMTPSERVER" { $newParms[$i] = "smtpServer=$smtpServer" }
      "SMTPFROMEMAIL" { $newParms[$i] = "smtpFromEmail=$smtpFromEmail" }
      "SMTPTOEMAIL" { $newParms[$i] = "smtpToEmail=$smtpToEmail" }
      "POSTLOG" { $newParms[$i] = "postLog=$postLog" }
    }
    $i++
  }

  if (test-path $propfileloc) {$newParms | Set-Content $propfileloc}
  else {$newParms | Add-Content $propfileloc}
  
  $txt_Info.AppendText("New configuration saved to " + $propfileloc)
  $txt_Info.AppendText("`n")
  #Scroll to the end of the textbox
  $txt_Info.SelectionStart = $txt_Info.TextLength;
  $txt_Info.ScrollToCaret()
  Start-Sleep -Seconds 1
} #saveConfig

# Display the Input form used for updating config/parms
function displayForm() {
  #$obj_tt = New-Object System.Windows.Forms.ToolTip $form.Container
  $obj_tt = New-Object System.Windows.Forms.ToolTip 
  $obj_tt.InitialDelay = 100     
  $obj_tt.ReshowDelay = 100 
  $obj_tt.IsBalloon = $true
  $obj_tt.Tag = "tooltip info"
  $obj_tt.ShowAlways = $true

  #Preload the presets array from 
  $presetsArray=@()
  $items = Get-ChildItem $hbpreloc -Recurse | Select-String -Pattern '"PresetName": '
  foreach ($item in $items) {
      $preset = ((($item.Line).Split(":"))[1]).replace('"','').replace(',','').trim()
      $presetsArray += $preset
  }
  $presetsArray = ($presetsArray | Sort-Object)

  $form                            = New-Object system.Windows.Forms.Form
  $form.Text                       = "VidMonHB"
  #$form.StartPosition              = 'CenterScreen'
  $Form.StartPosition              = 'Manual'
  $Form.Location                   = '0, 0'
  $form.ClientSize                 = [System.Drawing.Size]::new(1456,780)
  $form.FormBorderStyle            = $bs3
  $form.BackColor                  = $blue
  $form.TopMost                    = $false
  $form.AutoScroll                 = $true
  $form.AcceptButton               = $btn_execute
  $form.Add_KeyDown({if ($_.KeyCode -eq "Escape") { $script:exit = $true;$form.Dispose() } }) # if escape, exit
  $form.KeyPreview                 = $true
  $form.ShowInTaskbar              = $true
  
  $gbx_title                       = New-Object system.Windows.Forms.Groupbox
  $gbx_title.height                = 72
  $gbx_title.width                 = 1370
  $gbx_title.BackColor             = $blue
  $gbx_title.ForeColor             = $white
  $gbx_title.location              = New-Object System.Drawing.Point(48,10)

  $lbl_title                       = New-Object system.Windows.Forms.Label
  #$lbl_title.text                  = "VidMonHB"
  $lbl_title.BackColor             = $blue
  $lbl_title.AutoSize              = $true
  $lbl_title.width                 = 25
  $lbl_title.height                = 10
  $lbl_title.location              = New-Object System.Drawing.Point(12,15)
  $lbl_title.Font                  = $v16bi.font
  $lbl_title.ForeColor             = $yellow

  $pbx_image                       = new-object Windows.Forms.PictureBox
  $pbx_imageName                   = [system.drawing.image]::FromFile($vidMonHBImage) 
  $pbx_image.Image                 = $pbx_imageName
  $pbx_image.Width                 = 150
  $pbx_image.Height                = 40
  $pbx_image.location              = New-Object System.Drawing.Point(10,12)
  $pbx_image.AutoSize              = $true
  if (-not(Test-Path $vidMonHBImage)) {$lbl_title.text = "VidMonHB"}
  
  $lbl_heading                     = New-Object system.Windows.Forms.Label
  $lbl_heading.text                = "Handbrake Batch Converter"
  $lbl_heading.BackColor           = $blue
  $lbl_heading.AutoSize            = $true
  $lbl_heading.width               = 25
  $lbl_heading.height              = 10
  $lbl_heading.location            = New-Object System.Drawing.Point(342,9)
  $lbl_heading.Font                = $v30b.font
  $lbl_heading.ForeColor           = $yellow

  $lbl_version                     = New-Object system.Windows.Forms.Label
  $lbl_version.text                = "Version " + $version
  $lbl_version.BackColor           = $blue
  $lbl_version.AutoSize            = $true
  $lbl_version.width               = 25
  $lbl_version.height              = 10
  $lbl_version.location            = New-Object System.Drawing.Point(1154,20)
  $lbl_version.Font                = $v16bi.font
  $lbl_version.ForeColor           = $yellow

  $gbx_location                    = New-Object system.Windows.Forms.Groupbox
  $gbx_location.height             = 245
  $gbx_location.width              = 641
  $gbx_location.text               = "File Location Info"
  $gbx_location.Font               = $v9bi.font
  $gbx_location.ForeColor          = $white
  $gbx_location.location           = New-Object System.Drawing.Point(46,105)

  $gbx_hbInfo                      = New-Object system.Windows.Forms.Groupbox
  $gbx_hbInfo.height               = 245
  $gbx_hbInfo.width                = 701
  $gbx_hbInfo.text                 = "Handbrake Info"
  $gbx_hbInfo.Font                 = $v9bi.font
  $gbx_hbInfo.ForeColor            = $white
  $gbx_hbInfo.location             = New-Object System.Drawing.Point(719,105)

  $gbx_convInfo                    = New-Object system.Windows.Forms.Groupbox
  $gbx_convInfo.height             = 260
  $gbx_convInfo.width              = 641
  $gbx_convInfo.text               = "Conversion Options"
  $gbx_convInfo.Font               = $v9bi.font
  $gbx_convInfo.ForeColor          = $white
  $gbx_convInfo.location           = New-Object System.Drawing.Point(47,358)

  $gbx_postInfo                    = New-Object system.Windows.Forms.Groupbox
  $gbx_postInfo.height             = 260
  $gbx_postInfo.width              = 701
  $gbx_postInfo.text               = "Post Processing Info"
  $gbx_postInfo.Font               = $v9bi.font
  $gbx_postInfo.ForeColor          = $white
  $gbx_postInfo.location           = New-Object System.Drawing.Point(719,358)

  $lbl_in                          = New-Object system.Windows.Forms.Label
  $lbl_in.text                     = "Input"
  $lbl_in.AutoSize                 = $true
  $lbl_in.width                    = 25
  $lbl_in.height                   = 10
  $lbl_in.location                 = New-Object System.Drawing.Point(98,22)
  $lbl_in.Font                     = $v12.font
  $lbl_in.ForeColor                = $lblColor

  $txt_in                          = New-Object system.Windows.Forms.TextBox
  $txt_in.TabIndex                 = 1
  $txt_in.multiline                = $false
  $txt_in.text                     = $in
  $txt_in.width                    = 469
  $txt_in.height                   = 20
  $txt_in.location                 = New-Object System.Drawing.Point(150,20)
  $txt_in.Font                     = $v12.font
  $txt_in.ForeColor                = $blue
  $txt_in.Add_GotFocus({ Paint-FocusBorder $this
                         $txt_Info.text = ($txt_in.Tag) 
                         if (-not (chkPath $this)) {
                           $txt_info.text = "The specified file path is invalid."
                           $txt_info.backColor = $yellow
                           $txt_info.foreColor = $red
                          }
                      })
  $txt_in.Add_LostFocus({ Paint-FocusBorder $this
                          $script:in=$txt_in.text
                          chkAllPaths
                          $txt_info.backColor = $white
                          $txt_info.foreColor = $blue
                        })
  $txt_in.Add_MouseEnter({ Show-ToolTip $this })
  $txt_in.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_in.Tag = "This is the disk location where video files will be searched for.`nScript will also recurse subdirectories."

  $lbl_out                         = New-Object system.Windows.Forms.Label
  $lbl_out.text                    = "Output"
  $lbl_out.AutoSize                = $true
  $lbl_out.width                   = 25
  $lbl_out.height                  = 10
  $lbl_out.location                = New-Object System.Drawing.Point(86,58)
  $lbl_out.Font                    = $v12.font
  $lbl_out.ForeColor               = $lblColor

  $lbl_out2                         = New-Object system.Windows.Forms.Label
  $lbl_out2.text                    = "** Output will be written to input location **"
  $lbl_out2.AutoSize                = $true
  $lbl_out2.width                   = 25
  $lbl_out2.height                  = 10
  $lbl_out2.location                = New-Object System.Drawing.Point(150,57)
  $lbl_out2.Font                    = $v12bi.font
  $lbl_out2.ForeColor               = $cyan #$lblColor
  $lbl_out2.Visible                 = $false

  $txt_out                         = New-Object system.Windows.Forms.TextBox
  $txt_out.TabIndex                = 2
  $txt_out.multiline               = $false
  $txt_out.text                    = $out
  $txt_out.width                   = 470
  $txt_out.height                  = 20
  $txt_out.location                = New-Object System.Drawing.Point(150,55)
  $txt_out.Font                    = $v12.font
  $txt_out.ForeColor               = $blue
  $txt_out.Add_GotFocus({ Paint-FocusBorder $this
                          $txt_Info.text = ($txt_out.Tag)
                          if (-not (chkPath $this)) {
                            $txt_info.text = "The specified file path is invalid."
                            $txt_info.backColor = $yellow
                            $txt_info.foreColor = $red
                           }
                         })
  $txt_out.Add_LostFocus({ Paint-FocusBorder $this
                          $script:out=$txt_out.text
                          chkAllPaths
                          $txt_info.backColor = $white
                          $txt_info.foreColor = $blue
                        })
  $txt_out.Add_MouseEnter({ Show-ToolTip $this })
  $txt_out.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_out.Tag = "This is the disk location where converted video files will be written to."
  
  $lbl_logfilePath                 = New-Object system.Windows.Forms.Label
  $lbl_logfilePath.text            = "Log files"
  $lbl_logfilePath.AutoSize        = $true
  $lbl_logfilePath.width           = 25
  $lbl_logfilePath.height          = 10
  $lbl_logfilePath.location        = New-Object System.Drawing.Point(74,94)
  $lbl_logfilePath.Font            = $v12.font
  $lbl_logfilePath.ForeColor       = $lblColor

  $txt_logfilePath                 = New-Object system.Windows.Forms.TextBox
  $txt_logfilePath.TabIndex        = 3
  $txt_logfilePath.multiline       = $false
  $txt_logfilePath.text            = $logfilePath
  $txt_logfilePath.width           = 470
  $txt_logfilePath.height          = 20
  $txt_logfilePath.location        = New-Object System.Drawing.Point(150,92)
  $txt_logfilePath.Font            = $v12.font
  $txt_logfilePath.ForeColor       = $blue
  $txt_logfilePath.Add_GotFocus({ Paint-FocusBorder $this
                                  $txt_Info.text = ($txt_logfilePath.Tag)
                                  if (-not (chkPath $this)) {
                                    $txt_info.text = "The specified file path is invalid."
                                    $txt_info.backColor = $yellow
                                    $txt_info.foreColor = $red
                                   }
                               })
  $txt_logfilePath.Add_LostFocus({ Paint-FocusBorder $this
                                   $script:logfilePath=$txt_logfilePath.text
                                   chkAllPaths
                                   $txt_info.backColor = $white
                                   $txt_info.foreColor = $blue
                                })
  $txt_logfilePath.Add_MouseEnter({ Show-ToolTip $this })
  $txt_logfilePath.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_logfilePath.Tag = "This is the disk location where log files will be written to."

  $lbl_TVShowBasePath              = New-Object system.Windows.Forms.Label
  $lbl_TVShowBasePath.text         = "TV Base Path"
  $lbl_TVShowBasePath.AutoSize     = $true
  $lbl_TVShowBasePath.width        = 25
  $lbl_TVShowBasePath.height       = 10
  $lbl_TVShowBasePath.location     = New-Object System.Drawing.Point(34,157)
  $lbl_TVShowBasePath.Font         = $v12.font
  $lbl_TVShowBasePath.ForeColor    = $lblColor

  $txt_TVShowBasePath              = New-Object system.Windows.Forms.TextBox
  $txt_TVShowBasePath.TabIndex     = 4
  $txt_TVShowBasePath.multiline    = $false
  $txt_TVShowBasePath.text         = $TVShowBasePath
  $txt_TVShowBasePath.width        = 470
  $txt_TVShowBasePath.height       = 20
  $txt_TVShowBasePath.location     = New-Object System.Drawing.Point(150,157)
  $txt_TVShowBasePath.Font         = $v12.font
  $txt_TVShowBasePath.ForeColor    = $blue
  $txt_TVShowBasePath.Add_GotFocus({ Paint-FocusBorder $this
                                     $txt_Info.text = ($txt_TVShowBasePath.Tag) 
                                     if (-not (chkPath $this)) {
                                      $txt_info.text = "The specified file path is invalid."
                                      $txt_info.backColor = $yellow
                                      $txt_info.foreColor = $red
                                     }
                                  })
  $txt_TVShowBasePath.Add_LostFocus({ Paint-FocusBorder $this
                                      $script:TVShowBasePath=$txt_TVShowBasePath.text
                                      chkAllPaths
                                      $txt_info.backColor = $white
                                      $txt_info.foreColor = $blue
                                   })
  $txt_TVShowBasePath.Add_MouseEnter({ Show-ToolTip $this })
  $txt_TVShowBasePath.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_TVShowBasePath.Tag = "This is the 'Base' disk location where converted TV Show files will be written to."

  $lbl_movieBasePath               = New-Object system.Windows.Forms.Label
  $lbl_movieBasePath.text          = "Movie Base Path"
  $lbl_movieBasePath.AutoSize      = $true
  $lbl_movieBasePath.width         = 25
  $lbl_movieBasePath.height        = 10
  $lbl_movieBasePath.location      = New-Object System.Drawing.Point(5,193)
  $lbl_movieBasePath.Font          = $v12.font
  $lbl_movieBasePath.ForeColor     = $lblColor

  $txt_movieBasePath               = New-Object system.Windows.Forms.TextBox
  $txt_movieBasePath.TabIndex      = 5
  $txt_movieBasePath.multiline     = $false
  $txt_movieBasePath.text          = $movieBasePath
  $txt_movieBasePath.width         = 470
  $txt_movieBasePath.height        = 20
  $txt_movieBasePath.location      = New-Object System.Drawing.Point(150,191)
  $txt_movieBasePath.Font          = $v12.font
  $txt_movieBasePath.ForeColor     = $blue
  $txt_movieBasePath.Add_GotFocus({ Paint-FocusBorder $this
                                    $txt_Info.text = ($txt_movieBasePath.Tag) 
                                    if (-not (chkPath $this)) {
                                      $txt_info.text = "The specified file path is invalid."
                                      $txt_info.backColor = $yellow
                                      $txt_info.foreColor = $red
                                     }
                                  })
  $txt_movieBasePath.Add_LostFocus({ Paint-FocusBorder $this
                                     $script:movieBasePath=$txt_movieBasePath.text
                                     chkAllPaths
                                     $txt_info.backColor = $white
                                     $txt_info.foreColor = $blue
                                  })
  $txt_movieBasePath.Add_MouseEnter({ Show-ToolTip $this })
  $txt_movieBasePath.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_movieBasePath.Tag = "This is the 'Base' disk location where converted Movie files will be written to."

  $lbl_hbloc                       = New-Object system.Windows.Forms.Label
  $lbl_hbloc.text                  = "Handbrake Program"
  $lbl_hbloc.AutoSize              = $true
  $lbl_hbloc.width                 = 25
  $lbl_hbloc.height                = 10
  $lbl_hbloc.location              = New-Object System.Drawing.Point(12,25)
  $lbl_hbloc.Font                  = $v12.font
  $lbl_hbloc.ForeColor             = $lblColor

  $txt_hbloc                       = New-Object system.Windows.Forms.TextBox
  $txt_hbloc.TabIndex              = 6
  $txt_hbloc.multiline             = $false
  $txt_hbloc.text                  = $hbloc
  $txt_hbloc.width                 = 500
  $txt_hbloc.height                = 20
  $txt_hbloc.location              = New-Object System.Drawing.Point(186,25)
  $txt_hbloc.Font                  = $v12.font
  $txt_hbloc.ForeColor             = $blue
  $txt_hbloc.Add_GotFocus({ Paint-FocusBorder $this
                            $txt_Info.text = ($txt_hbloc.Tag) 
                            if (-not (chkPath $this)) {
                              $txt_info.text = "The specified file path is invalid."
                              $txt_info.backColor = $yellow
                              $txt_info.foreColor = $red
                             }
                         })
  $txt_hbloc.Add_LostFocus({ Paint-FocusBorder $this
                             $script:hbloc=$txt_hbloc.text
                             chkAllPaths
                             $txt_info.backColor = $white
                             $txt_info.foreColor = $blue
                          })
  $txt_hbloc.Add_MouseEnter({ Show-ToolTip $this })
  $txt_hbloc.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_hbloc.Tag = "This is the disk location where the HandbrakeCLI.exe program can be found."
  
  $lbl_hbpreloc                    = New-Object system.Windows.Forms.Label
  $lbl_hbpreloc.text               = "Presets Location"
  $lbl_hbpreloc.AutoSize           = $true
  $lbl_hbpreloc.width              = 25
  $lbl_hbpreloc.height             = 10
  $lbl_hbpreloc.location           = New-Object System.Drawing.Point(44,60)
  $lbl_hbpreloc.Font               = $v12.font
  $lbl_hbpreloc.ForeColor          = $lblColor

  $txt_hbpreloc                    = New-Object system.Windows.Forms.TextBox
  $txt_hbpreloc.TabIndex           = 7
  $txt_hbpreloc.multiline          = $false
  $txt_hbpreloc.text               = $hbpreloc
  $txt_hbpreloc.width              = 500
  $txt_hbpreloc.height             = 20
  $txt_hbpreloc.location           = New-Object System.Drawing.Point(186,59)
  $txt_hbpreloc.Font               = $v12.font
  $txt_hbpreloc.ForeColor          = $blue
  $txt_hbpreloc.Add_GotFocus({ Paint-FocusBorder $this
                               $txt_Info.text = ($txt_hbpreloc.Tag) 
                               if (-not (chkPath $this)) {
                                $txt_info.text = "The specified file path is invalid."
                                $txt_info.backColor = $yellow
                                $txt_info.foreColor = $red
                               }
                            })
  $txt_hbpreloc.Add_LostFocus({ Paint-FocusBorder $this
                                $script:hbpreloc=$txt_hbpreloc.text
                                chkAllPaths
                                $txt_info.backColor = $white
                                $txt_info.foreColor = $blue
                             })
  $txt_hbpreloc.Add_MouseEnter({ Show-ToolTip $this })
  $txt_hbpreloc.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_hbpreloc.Tag = "This is the disk location where the presets.json file can be found."

  $lbl_hbopts                      = New-Object system.Windows.Forms.Label
  $lbl_hbopts.text                 = "Addl HB Options"
  $lbl_hbopts.AutoSize             = $true
  $lbl_hbopts.width                = 25
  $lbl_hbopts.height               = 10
  $lbl_hbopts.location             = New-Object System.Drawing.Point(44,94)
  $lbl_hbopts.Font                 = $v12.font
  $lbl_hbopts.ForeColor            = $lblColor

  $txt_hbopts                      = New-Object system.Windows.Forms.TextBox
  $txt_hbopts.TabIndex             = 8
  $txt_hbopts.multiline            = $false
  $txt_hbopts.text                 = $hbopts
  $txt_hbopts.width                = 500
  $txt_hbopts.height               = 20
  $txt_hbopts.location             = New-Object System.Drawing.Point(186,92)
  $txt_hbopts.Font                 = $v12.font
  $txt_hbopts.ForeColor            = $blue
  $txt_hbopts.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($txt_hbopts.Tag) })
  $txt_hbopts.Add_LostFocus({ Paint-FocusBorder $this; $script:hbopts=$txt_hbopts.text })
  $txt_hbopts.Add_MouseEnter({ Show-ToolTip $this })
  $txt_hbopts.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_hbopts.Tag = "Additional options specified here will be appended to the Handbrake command."

  $lbl_tvPreset                    = New-Object system.Windows.Forms.Label
  $lbl_tvPreset.text               = "TV Preset Name"
  $lbl_tvPreset.AutoSize           = $true
  $lbl_tvPreset.width              = 25
  $lbl_tvPreset.height             = 10
  $lbl_tvPreset.location           = New-Object System.Drawing.Point(44,158)
  $lbl_tvPreset.Font               = $v12.font
  $lbl_tvPreset.ForeColor          = $lblColor

  $cbx_tvPreset                    = New-Object system.Windows.Forms.ComboBox
  $cbx_tvPreset.TabIndex           = 9
  $cbx_tvPreset.text               = $tvPreset
  $cbx_tvPreset.width              = 340
  $cbx_tvPreset.height             = 20
  $cbx_tvPreset.location           = New-Object System.Drawing.Point(186,158)
  $cbx_tvPreset.Font               = $ss12.font
  $cbx_tvPreset.ForeColor          = $blue
  $cbx_tvPreset.AutoCompleteMode   = 3 #SuggestAppend
  $currentText = ($cbx_tvPreset).Text
  $cbx_tvPreset.Items.Clear()
  $cbx_tvPreset.Text = ""
  $presetsArray | ForEach-Object {[void] $cbx_tvPreset.Items.Add($_)}
  $cbx_tvPreset.Text               = $currentText
  $length = $presetsArray | Measure-Object -Maximum -Property Length | Select-Object maximum
  $cbx_tvPreset.DropDownWidth = 240 * $length.Maximum / 25
  $cbx_tvPreset.Add_GotFocus({ Paint-FocusBorder $this
                               $txt_Info.text = ($cbx_tvPreset.Tag) 
                               if (-not (chkPreset $this)) {
                                $txt_info.text = "A preset named " + $cbx_tvPreset.Text + 
                                                 " could not be found in preset file " + $txt_hbpreloc.text
                                $txt_info.backColor = $yellow
                                $txt_info.foreColor = $red
                               }
                            })
  $cbx_tvPreset.Add_LostFocus({ Paint-FocusBorder $this
                                $script:tvPreset=$cbx_tvPreset.text 
                                chkPreset $this
                                $txt_info.backColor = $white
                                $txt_info.foreColor = $blue
                              })
  $cbx_tvPreset.Add_MouseEnter({ Show-ToolTip $this })
  $cbx_tvPreset.Add_MouseLeave({ $obj_tt.Hide($form) })
  $cbx_tvPreset.Tag = "The preset name specified here will be used for converting TV show files." 

  $lbl_moviePreset                 = New-Object system.Windows.Forms.Label
  $lbl_moviePreset.text            = "Movie Preset Name"
  $lbl_moviePreset.AutoSize        = $true
  $lbl_moviePreset.width           = 25
  $lbl_moviePreset.height          = 10
  $lbl_moviePreset.location        = New-Object System.Drawing.Point(18,195)
  $lbl_moviePreset.Font            = $v12.font
  $lbl_moviePreset.ForeColor       = $lblColor

  $cbx_moviePreset                 = New-Object system.Windows.Forms.ComboBox
  $cbx_moviePreset.TabIndex        = 10
  $cbx_moviePreset.text            = $moviePreset
  $cbx_moviePreset.width           = 340
  $cbx_moviePreset.height          = 20
  $cbx_moviePreset.location        = New-Object System.Drawing.Point(186,194)
  $cbx_moviePreset.Font            = $ss12.font
  $cbx_moviePreset.ForeColor       = $blue
  $cbx_moviePreset.AutoCompleteMode= 3 #SuggestAppend
  $currentText = ($cbx_moviePreset).Text
  $cbx_moviePreset.Items.Clear()
  $cbx_moviePreset.Text = ""
  $presetsArray | ForEach-Object {[void] $cbx_moviePreset.Items.Add($_)}
  $cbx_moviePreset.Text               = $currentText
  $length = $presetsArray | Measure-Object -Maximum -Property Length | Select-Object maximum
  $cbx_moviePreset.DropDownWidth = 240 * $length.Maximum / 25
  $cbx_moviePreset.Add_GotFocus({ Paint-FocusBorder $this
                                  $txt_Info.text = ($cbx_moviePreset.Tag) 
                                  if (-not (chkPreset $this)) {
                                    $txt_info.text = "A preset named " + $cbx_moviePreset.Text + 
                                                     " could not be found in preset file " + $txt_hbpreloc.text
                                    $txt_info.backColor = $yellow
                                    $txt_info.foreColor = $red
                                   }
                                })
  $cbx_moviePreset.Add_LostFocus({ Paint-FocusBorder $this
                                   $script:moviePreset=$cbx_moviePreset.text 
                                   chkPreset $this
                                   $txt_info.backColor = $white
                                   $txt_info.foreColor = $blue
                                })
  $cbx_moviePreset.Add_MouseEnter({ Show-ToolTip $this })
  $cbx_moviePreset.Add_MouseLeave({ $obj_tt.Hide($form) })
  $cbx_moviePreset.Tag = "The preset name specified here will be used for converting Movie files." 

  $lbl_vidtypes                    = New-Object system.Windows.Forms.Label
  $lbl_vidtypes.text               = "Input Video Types"
  $lbl_vidtypes.AutoSize           = $true
  $lbl_vidtypes.width              = 25
  $lbl_vidtypes.height             = 10
  $lbl_vidtypes.location           = New-Object System.Drawing.Point(23,32)
  $lbl_vidtypes.Font               = $v12.font
  $lbl_vidtypes.ForeColor          = $lblColor

  $txt_vidtypes                    = New-Object system.Windows.Forms.TextBox
  $txt_vidtypes.TabIndex           = 11
  $txt_vidtypes.multiline          = $false
  $txt_vidtypes.text               = $vidTypes
  $txt_vidtypes.width              = 427
  $txt_vidtypes.height             = 20
  $txt_vidtypes.location           = New-Object System.Drawing.Point(186,29)
  $txt_vidtypes.Font               = $v12.font
  $txt_vidtypes.ForeColor          = $blue
  $txt_vidtypes.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($txt_vidtypes.Tag) })
  $txt_vidtypes.Add_LostFocus({ Paint-FocusBorder $this; $script:vidtypes=$txt_vidtypes.text })
  $txt_vidtypes.Add_MouseEnter({ Show-ToolTip $this })
  $txt_vidtypes.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_vidtypes.Tag = "Specify the video extensions to search for " +
                      " (i.e. MPG,MP2,MPEG,MPE,MPV,OGG,M4P,M4V,AVI,WMV,MOV,QT,FLV,SWF,WEBM)."

  $lbl_limit                       = New-Object system.Windows.Forms.Label
  $lbl_limit.text                  = "# of files to process"
  $lbl_limit.AutoSize              = $true
  $lbl_limit.width                 = 25
  $lbl_limit.height                = 10
  $lbl_limit.location              = New-Object System.Drawing.Point(9,67)
  $lbl_limit.Font                  = $v12.font
  $lbl_limit.ForeColor             = $lblColor

  $txt_limit                       = New-Object system.Windows.Forms.TextBox
  $txt_limit.TabIndex              = 12
  $txt_limit.multiline             = $false
  $txt_limit.text                  = $limit
  $txt_limit.width                 = 50
  $txt_limit.height                = 20
  $txt_limit.location              = New-Object System.Drawing.Point(186,64)
  $txt_limit.Font                  = $v12.font
  $txt_limit.ForeColor             = $blue
  $txt_limit.MaxLength             = 4
  $txt_limit.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($txt_limit.Tag) })
  $txt_limit.Add_LostFocus({ Paint-FocusBorder $this; $script:limit=$txt_limit.text })
  $txt_limit.Add_MouseEnter({ Show-ToolTip $this })
  $txt_limit.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_limit.Tag = "Specify the max # of files to convert."

  $lbl_ParallelProcMax             = New-Object system.Windows.Forms.Label
  $lbl_ParallelProcMax.text        = "Parallel processes"
  $lbl_ParallelProcMax.AutoSize    = $true
  $lbl_ParallelProcMax.width       = 25
  $lbl_ParallelProcMax.height      = 10
  $lbl_ParallelProcMax.location    = New-Object System.Drawing.Point(25,99)
  $lbl_ParallelProcMax.Font        = $v12.font
  $lbl_ParallelProcMax.ForeColor   = $lblColor

  $cbx_ParallelProcMax             = New-Object system.Windows.Forms.ComboBox
  $cbx_ParallelProcMax.TabIndex    = 13
  $cbx_ParallelProcMax.text        = $ParallelProcMax
  $cbx_ParallelProcMax.width       = 50
  $cbx_ParallelProcMax.height      = 20
  $cbx_ParallelProcMax.AutoCompleteMode= 3 #SuggestAppend
  $cbx_ParallelProcMax.DropDownStyle = 2 #DropDownList
  $currentText = ($cbx_ParallelProcMax).Text
  $cbx_ParallelProcMax.Items.Clear()
  $cbx_ParallelProcMax.Text = ""
  @('1','2','3','4','5','6','7','8','9','10') | ForEach-Object {[void] $cbx_ParallelProcMax.Items.Add($_)}
  $cbx_ParallelProcMax.Text        = $currentText
  $cbx_ParallelProcMax.location    = New-Object System.Drawing.Point(184,98)
  $cbx_ParallelProcMax.Font        = $v12.font
  $cbx_ParallelProcMax.ForeColor   = $blue
  $cbx_ParallelProcMax.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($cbx_ParallelProcMax.Tag) })
  $cbx_ParallelProcMax.Add_LostFocus({ Paint-FocusBorder $this; $script:ParallelProcMax=$cbx_ParallelProcMax.text; })
  $cbx_ParallelProcMax.Add_SelectedIndexChanged({ if ($cbx_ParallelProcMax.SelectedItem -gt 1) 
                                                {$lbl_ParallelProcMsg.Text = "Parallel Processing Enabled"} 
                                           else {$lbl_ParallelProcMsg.Text = ""} })
#  $cbx_ParallelProcMax.Add_MouseEnter({ Show-ToolTip $this })
#  $cbx_ParallelProcMax.Add_MouseLeave({ $obj_tt.Hide($form) })
  $cbx_ParallelProcMax.Tag = "Specify the max # of threads.`nWarning - Make sure your" +
                             " machine can handle multiple processes before selecting higher values."
                                           
  $lbl_ParallelProcMsg             = New-Object system.Windows.Forms.Label
  $lbl_ParallelProcMsg.text        = ""
  $lbl_ParallelProcMsg.AutoSize    = $true
  $lbl_ParallelProcMsg.width       = 25
  $lbl_ParallelProcMsg.height      = 10
  $lbl_ParallelProcMsg.location    = New-Object System.Drawing.Point(240,99)
  $lbl_ParallelProcMsg.Font        = $v12b.font
  $lbl_ParallelProcMsg.ForeColor   = $lblColor

  $gbx_outSameAsIn                 = New-Object system.Windows.Forms.Groupbox
  $gbx_outSameAsIn.TabIndex        = 14
  $gbx_outSameAsIn.TabStop         = $true
  $gbx_outSameAsIn.height          = 76
  $gbx_outSameAsIn.width           = 179
  $gbx_outSameAsIn.location        = New-Object System.Drawing.Point(12,135)

  $lbl_outSameAsIn                 = New-Object system.Windows.Forms.Label
  $lbl_outSameAsIn.text            = "Output dest same"
  $lbl_outSameAsIn.AutoSize        = $true
  $lbl_outSameAsIn.width           = 25
  $lbl_outSameAsIn.height          = 10
  $lbl_outSameAsIn.location        = New-Object System.Drawing.Point(15,16)
  $lbl_outSameAsIn.Font            = $ss12.font
  $lbl_outSameAsIn.ForeColor       = $lblColor

  $lbl_outSameAsIn2                = New-Object system.Windows.Forms.Label
  $lbl_outSameAsIn2.text           = "as input location"
  $lbl_outSameAsIn2.AutoSize       = $true
  $lbl_outSameAsIn2.width          = 25
  $lbl_outSameAsIn2.height         = 10
  $lbl_outSameAsIn2.location       = New-Object System.Drawing.Point(19,38)
  $lbl_outSameAsIn2.Font           = $ss12.font
  $lbl_outSameAsIn2.ForeColor      = $lblColor

  $cbx_outSameAsIn                 = New-Object system.Windows.Forms.CheckBox
  $cbx_outSameAsIn.Checked         = $outSameAsIn
  $cbx_outSameAsIn.AutoSize        = $false
  $cbx_outSameAsIn.width           = 20
  $cbx_outSameAsIn.height          = 17
  $cbx_outSameAsIn.location        = New-Object System.Drawing.Point(152,26)
  $cbx_outSameAsIn.Font            = $ss12.font
  $cbx_outSameAsIn.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($cbx_outSameAsIn.Tag) })
  $cbx_outSameAsIn.Add_LostFocus({ Paint-FocusBorder $this; $script:outSameAsIn=$cbx_outSameAsIn.Checked })
  $cbx_outSameAsIn.Add_CheckStateChanged({ 
    if ($cbx_outSameAsIn.Checked -eq $true) {$txt_out.Visible=$false; $lbl_out2.Visible=$true }
    else {$txt_out.Visible=$true; $lbl_out2.Visible=$false}
  })
  $cbx_outSameAsIn.Add_MouseEnter({ Show-ToolTip $this })
  $cbx_outSameAsIn.Add_MouseLeave({ $obj_tt.Hide($form) })
  $cbx_outSameAsIn.Tag = "Overrides the output destination folder.`n" + 
                         "Converted files will be written to the same folder where the original was found."

  $gbx_movefiles                   = New-Object system.Windows.Forms.Groupbox
  $gbx_movefiles.TabIndex          = 15
  $gbx_movefiles.TabStop           = $true
  $gbx_movefiles.height            = 76
  $gbx_movefiles.width             = 168
  $gbx_movefiles.location          = New-Object System.Drawing.Point(224,135)

  $lbl_movefiles                   = New-Object system.Windows.Forms.Label
  $lbl_movefiles.text              = "Move files after"
  $lbl_movefiles.AutoSize          = $true
  $lbl_movefiles.width             = 25
  $lbl_movefiles.height            = 10
  $lbl_movefiles.location          = New-Object System.Drawing.Point(16,17)
  $lbl_movefiles.Font              = $ss12.font
  $lbl_movefiles.ForeColor         = $lblColor

  $lbl_movefiles2                  = New-Object system.Windows.Forms.Label
  $lbl_movefiles2.text             = "conversion"
  $lbl_movefiles2.AutoSize         = $true
  $lbl_movefiles2.width            = 25
  $lbl_movefiles2.height           = 10
  $lbl_movefiles2.location         = New-Object System.Drawing.Point(32,38)
  $lbl_movefiles2.Font             = $ss12.font
  $lbl_movefiles2.ForeColor        = $lblColor

  $cbx_movefiles                   = New-Object system.Windows.Forms.CheckBox
  $cbx_movefiles.Checked           = $movefiles
  $cbx_movefiles.AutoSize          = $false
  $cbx_movefiles.width             = 20
  $cbx_movefiles.height            = 17
  $cbx_movefiles.location          = New-Object System.Drawing.Point(138,26)
  $cbx_movefiles.Font              = $ss12.font
  $cbx_movefiles.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($cbx_movefiles.Tag) })
  $cbx_movefiles.Add_LostFocus({ Paint-FocusBorder $this; $script:movefiles=$cbx_movefiles.Checked })
  $cbx_movefiles.Add_MouseEnter({ Show-ToolTip $this })
  $cbx_movefiles.Add_MouseLeave({ $obj_tt.Hide($form) })
  $cbx_movefiles.Tag = "Move converted file(s) to final TV or Movie destination." +
                       "`nTV Shows will be moved into folder names by show title" +
                       "`nIf the year is present in the movie title, movies will be moved into folder names " + 
                       "by year. Otherwise they will be moved into the Movie base folder."

  $gbx_delAfterConv                = New-Object system.Windows.Forms.Groupbox
  $gbx_delAfterConv.TabIndex       = 16
  $gbx_delAfterConv.TabStop        = $true
  $gbx_delAfterConv.height         = 76
  $gbx_delAfterConv.width          = 200
  $gbx_delAfterConv.location       = New-Object System.Drawing.Point(420,135)

  $lbl_delAfterConv                = New-Object system.Windows.Forms.Label
  $lbl_delAfterConv.text           = " Delete`nOriginal"
  $lbl_delAfterConv.AutoSize       = $true
  $lbl_delAfterConv.width          = 25
  $lbl_delAfterConv.height         = 10
  $lbl_delAfterConv.location       = New-Object System.Drawing.Point(9,15)
  $lbl_delAfterConv.Font           = $ss12.font
  $lbl_delAfterConv.ForeColor      = $lblColor

  $cbx_delAfterConv                = New-Object system.Windows.Forms.ComboBox
  $cbx_delAfterConv.Text           = $delAfterConv
  $cbx_delAfterConv.TabIndex       = 17
  $cbx_delAfterConv.width          = 100
  $cbx_delAfterConv.height         = 20
  $currentText                     = ($cbx_delAfterConv).Text
  $cbx_delAfterConv.Items.Clear()
  $cbx_delAfterConv.Text           = ""
  @('Maintain','Delete','Recycle') | ForEach-Object {[void] $cbx_delAfterConv.Items.Add($_)}
  $cbx_delAfterConv.Text           = $currentText
  $cbx_delAfterConv.location       = New-Object System.Drawing.Point(82,24)
  $cbx_delAfterConv.Font           = $ss12.font
  $cbx_delAfterConv.ForeColor      = $blue
  $cbx_delAfterConv.DropDownStyle  = 2 #DropDownList
  $cbx_delAfterConv.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($cbx_delAfterConv.Tag) })
  $cbx_delAfterConv.Add_LostFocus({ Paint-FocusBorder $this; $script:delAfterConv=$cbx_delAfterConv.text })
  $cbx_delAfterConv.Add_MouseEnter({ Show-ToolTip $this })
  $cbx_delAfterConv.Add_MouseLeave({ $obj_tt.Hide($form) })
  $cbx_delAfterConv.Tag = "Maintain, Delete or Recycle the original file after conversion." +
                          "`nNOTE - Recycle only works on local files. Files on network drives will be deleted!"

  <#
  $cbx_delAfterConv                = New-Object system.Windows.Forms.CheckBox
  $cbx_delAfterConv.Checked        = $delAfterConv
  $cbx_delAfterConv.AutoSize       = $false
  $cbx_delAfterConv.width          = 20
  $cbx_delAfterConv.height         = 17
  $cbx_delAfterConv.location       = New-Object System.Drawing.Point(122,28)
  $cbx_delAfterConv.Font           = $ss12.font
  $cbx_delAfterConv.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($cbx_delAfterConv.Tag) })
  $cbx_delAfterConv.Add_LostFocus({ Paint-FocusBorder $this; $script:delAfterConv=$cbx_delAfterConv.Checked })
  $cbx_delAfterConv.Add_MouseEnter({ Show-ToolTip $this })
  $cbx_delAfterConv.Add_MouseLeave({ $obj_tt.Hide($form) })
  $cbx_delAfterConv.Tag = "Delete the original file after conversion."
  #>

  $lbl_postExecCmd                 = New-Object system.Windows.Forms.Label
  $lbl_postExecCmd.text            = "Post execution cmd"
  $lbl_postExecCmd.AutoSize        = $true
  $lbl_postExecCmd.width           = 25
  $lbl_postExecCmd.height          = 10
  $lbl_postExecCmd.location        = New-Object System.Drawing.Point(17,32)
  $lbl_postExecCmd.Font            = $v12.font
  $lbl_postExecCmd.ForeColor       = $lblColor

  $txt_postExecCmd                 = New-Object system.Windows.Forms.TextBox
  $txt_postExecCmd.Text            = $postExecCmd
  $txt_postExecCmd.TabIndex        = 18
  $txt_postExecCmd.multiline       = $false
  $txt_postExecCmd.width           = 500
  $txt_postExecCmd.height          = 20
  $txt_postExecCmd.location        = New-Object System.Drawing.Point(186,26)
  $txt_postExecCmd.Font            = $v12.font
  $txt_postExecCmd.ForeColor       = $blue
  $txt_postExecCmd.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($txt_postExecCmd.Tag) })
  $txt_postExecCmd.Add_LostFocus({ Paint-FocusBorder $this; $script:postExecCmd=$txt_postExecCmd.text })
  $txt_postExecCmd.Add_MouseEnter({ Show-ToolTip $this })
  $txt_postExecCmd.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_postExecCmd.Tag = "Command to run after conversions complete."

  $lbl_postExecArgs                = New-Object system.Windows.Forms.Label
  $lbl_postExecArgs.text           = "Post execution args"
  $lbl_postExecArgs.AutoSize       = $true
  $lbl_postExecArgs.width          = 25
  $lbl_postExecArgs.height         = 10
  $lbl_postExecArgs.location       = New-Object System.Drawing.Point(17,67)
  $lbl_postExecArgs.Font           = $v12.font
  $lbl_postExecArgs.ForeColor      = $lblColor

  $txt_postExecArgs                = New-Object system.Windows.Forms.TextBox
  $txt_postExecArgs.Text           = $postExecArgs
  $txt_postExecArgs.TabIndex       = 19
  $txt_postExecArgs.multiline      = $false
  $txt_postExecArgs.width          = 500
  $txt_postExecArgs.height         = 20
  $txt_postExecArgs.location       = New-Object System.Drawing.Point(186,63)
  $txt_postExecArgs.Font           = $v12.font
  $txt_postExecArgs.ForeColor      = $blue
  $txt_postExecArgs.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($txt_postExecArgs.Tag) })
  $txt_postExecArgs.Add_LostFocus({ Paint-FocusBorder $this; $script:postExecArgs=$txt_postExecArgs.text })
  $txt_postExecArgs.Add_MouseEnter({ Show-ToolTip $this })
  $txt_postExecArgs.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_postExecArgs.Tag = "Arguments to be passed to the post execution command."

  $gbx_postNotification            = New-Object system.Windows.Forms.Groupbox
  $gbx_postNotification.height     = 145
  $gbx_postNotification.width      = 420
  #$gbx_postNotification.text       = "File Location Info"
  $gbx_postNotification.Font       = $v9bi.font
  $gbx_postNotification.ForeColor  = $white
  $gbx_postNotification.location   = New-Object System.Drawing.Point(35,92)

  $lbl_postNotify                  = New-Object system.Windows.Forms.Label
  $lbl_postNotify.text             = "Post notification"
  $lbl_postNotify.AutoSize         = $true
  $lbl_postNotify.width            = 25
  $lbl_postNotify.height           = 10
  $lbl_postNotify.location         = New-Object System.Drawing.Point(44,107)
  $lbl_postNotify.Font             = $v12.font
  $lbl_postNotify.ForeColor        = $lblColor

  $cbx_postNotify                  = New-Object system.Windows.Forms.ComboBox
  $cbx_postNotify.Text             = $postNotify
  $cbx_postNotify.TabIndex         = 20
#  $cbx_postNotify.text             = "None"
  $cbx_postNotify.width            = 80
  $cbx_postNotify.height           = 20
  $currentText = ($cbx_postNotify).Text
  $cbx_postNotify.Items.Clear()
  $cbx_postNotify.Text = ""
  @('None','All','Error') | ForEach-Object {[void] $cbx_postNotify.Items.Add($_)}
  $cbx_postNotify.Text             = $currentText
  $cbx_postNotify.location         = New-Object System.Drawing.Point(186,104)
  $cbx_postNotify.Font             = $ss12.font
  $cbx_postNotify.ForeColor        = $blue
  $cbx_postNotify.DropDownStyle    = 2 #DropDownList
  $cbx_postNotify.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($cbx_postNotify.Tag) })
  $cbx_postNotify.Add_LostFocus({ Paint-FocusBorder $this; $script:postNotify=$cbx_postNotify.text })
  $cbx_postNotify.Add_SelectedIndexChanged({
    if ($cbx_postNotify.SelectedItem -eq "None") {  
      $txt_smtpServer.Enabled          = $false
      $txt_smtpFromEmail.Enabled       = $false
      $txt_smtpToEmail.Enabled         = $false
    }
    else {
      $txt_smtpServer.Enabled          = $true
      $txt_smtpFromEmail.Enabled       = $true
      $txt_smtpToEmail.Enabled         = $true
      }
    })
  $cbx_postNotify.Add_MouseEnter({ Show-ToolTip $this })
  $cbx_postNotify.Add_MouseLeave({ $obj_tt.Hide($form) })
  $cbx_postNotify.Tag = "Type of notification to send after conversion completes." +
                        "`nAll = Send log after conversions have completed." +
                        "`nError = Only send logs if an error occurs."

  $lbl_smtpServer                  = New-Object system.Windows.Forms.Label
  $lbl_smtpServer.text             = "SMTP Server"
  $lbl_smtpServer.AutoSize         = $true
  $lbl_smtpServer.width            = 25
  $lbl_smtpServer.height           = 10
  $lbl_smtpServer.location         = New-Object System.Drawing.Point(69,140)
  $lbl_smtpServer.Font             = $v12.font
  $lbl_smtpServer.ForeColor        = $lblColor

  $txt_smtpServer                  = New-Object system.Windows.Forms.TextBox
  $txt_smtpServer.Text             = $smtpServer
  $txt_smtpServer.TabIndex         = 21
  $txt_smtpServer.multiline        = $false
  $txt_smtpServer.width            = 250
  $txt_smtpServer.height           = 20
  $txt_smtpServer.location         = New-Object System.Drawing.Point(186,137)
  $txt_smtpServer.Font             = $v12.font
  $txt_smtpServer.ForeColor        = $blue
  $txt_smtpServer.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($txt_smtpServer.Tag) })
  $txt_smtpServer.Add_LostFocus({ Paint-FocusBorder $this; $script:smtpServer=$txt_smtpServer.Text })
  $txt_smtpServer.Add_MouseEnter({ Show-ToolTip $this })
  $txt_smtpServer.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_smtpServer.Tag = "SMTP Server Information."

  $lbl_smtpFromEmail               = New-Object system.Windows.Forms.Label
  $lbl_smtpFromEmail.text          = "From email"
  $lbl_smtpFromEmail.AutoSize      = $true
  $lbl_smtpFromEmail.width         = 25
  $lbl_smtpFromEmail.height        = 10
  $lbl_smtpFromEmail.location      = New-Object System.Drawing.Point(80,171)
  $lbl_smtpFromEmail.Font          = $v12.font
  $lbl_smtpFromEmail.ForeColor     = $lblColor

  $txt_smtpFromEmail               = New-Object system.Windows.Forms.TextBox
  $txt_smtpFromEmail.Text          = $smtpFromEmail
  $txt_smtpFromEmail.TabIndex      = 22
  $txt_smtpFromEmail.multiline     = $false
  $txt_smtpFromEmail.width         = 250
  $txt_smtpFromEmail.height        = 20
  $txt_smtpFromEmail.location      = New-Object System.Drawing.Point(186,169)
  $txt_smtpFromEmail.Font          = $v12.font
  $txt_smtpFromEmail.ForeColor     = $blue
  $txt_smtpFromEmail.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($txt_smtpFromEmail.Tag) })
  $txt_smtpFromEmail.Add_LostFocus({ Paint-FocusBorder $this; $script:smtpFromEmail=$txt_smtpFromEmail.Text })
  $txt_smtpFromEmail.Add_MouseEnter({ Show-ToolTip $this })
  $txt_smtpFromEmail.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_smtpFromEmail.Tag = "Email address to send post notification from."

  $lbl_smtpToEmail                 = New-Object system.Windows.Forms.Label
  $lbl_smtpToEmail.text            = "To email"
  $lbl_smtpToEmail.AutoSize        = $true
  $lbl_smtpToEmail.width           = 25
  $lbl_smtpToEmail.height          = 10
  $lbl_smtpToEmail.location        = New-Object System.Drawing.Point(104,206)
  $lbl_smtpToEmail.Font            = $v12.font
  $lbl_smtpToEmail.ForeColor       = $lblColor

  $txt_smtpToEmail                 = New-Object system.Windows.Forms.TextBox
  $txt_smtpToEmail.Text            = $smtpToEmail
  $txt_smtpToEmail.TabIndex        = 23
  $txt_smtpToEmail.multiline       = $false
  $txt_smtpToEmail.width           = 250
  $txt_smtpToEmail.height          = 20
  $txt_smtpToEmail.location        = New-Object System.Drawing.Point(186,202)
  $txt_smtpToEmail.Font            = $v12.font
  $txt_smtpToEmail.ForeColor       = $blue
  $txt_smtpToEmail.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($txt_smtpToEmail.Tag) })
  $txt_smtpToEmail.Add_LostFocus({ Paint-FocusBorder $this; $script:smtpToEmail=$txt_smtpToEmail.Text})
  $txt_smtpToEmail.Add_MouseEnter({ Show-ToolTip $this })
  $txt_smtpToEmail.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_smtpToEmail.Tag = "Email address(s) to send post notification to."

  $lbl_postLog                     = New-Object system.Windows.Forms.Label
  $lbl_postLog.text                = "Post open log"
  $lbl_postLog.AutoSize            = $true
  $lbl_postLog.width               = 25
  $lbl_postLog.height              = 10
  $lbl_postLog.location            = New-Object System.Drawing.Point(475,107)
  $lbl_postLog.Font                = $v12.font
  $lbl_postLog.ForeColor           = $lblColor

  $cbx_postLog                     = New-Object system.Windows.Forms.ComboBox
  $cbx_postLog.Text                = $postLog
  $cbx_postLog.TabIndex            = 24
  $cbx_postLog.width               = 80
  $cbx_postLog.height              = 20
  $currentText                     = ($cbx_postLog).Text
  $cbx_postLog.Items.Clear()
  $cbx_postLog.Text                = ""
  @('Always','Error','Never') | ForEach-Object {[void] $cbx_postLog.Items.Add($_)}
  $cbx_postLog.Text                = $currentText
  $cbx_postLog.location            = New-Object System.Drawing.Point(600,104)
  $cbx_postLog.Font                = $ss12.font
  $cbx_postLog.ForeColor           = $blue
  $cbx_postLog.DropDownStyle       = 2 #DropDownList
  $cbx_postLog.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($cbx_postLog.Tag) })
  $cbx_postLog.Add_LostFocus({ Paint-FocusBorder $this; $script:postLog=$cbx_postLog.text })
  $cbx_postLog.Add_MouseEnter({ Show-ToolTip $this })
  $cbx_postLog.Add_MouseLeave({ $obj_tt.Hide($form) })
  $cbx_postLog.Tag = "Should the log automatically be opened after completion." +
                        "`nAlways = Always open log file." +
                        "`n   Error = Only open log if an error occurs." +
                        "`n  Never = Never open the log file."


  $lbl_repeatCtr                   = New-Object system.Windows.Forms.Label
  $lbl_repeatCtr.text              = "Repeat Runs"
  $lbl_repeatCtr.AutoSize          = $true
  $lbl_repeatCtr.width             = 24
  $lbl_repeatCtr.height            = 10
  $lbl_repeatCtr.location          = New-Object System.Drawing.Point(485,142)
  $lbl_repeatCtr.Font              = $v12.font
  $lbl_repeatCtr.ForeColor         = $lblColor

  $txt_repeatCtr                   = New-Object system.Windows.Forms.TextBox
  $txt_repeatCtr.MaxLength         = 4
  $txt_repeatCtr.Text              = $repeatCtr
  $txt_repeatCtr.TabIndex          = 25
  $txt_repeatCtr.multiline         = $false
  $txt_repeatCtr.width             = 60
  $txt_repeatCtr.height            = 20
  $txt_repeatCtr.location          = New-Object System.Drawing.Point(600,142)
  $txt_repeatCtr.Font              = $v12.font
  $txt_repeatCtr.ForeColor         = $blue
  $txt_repeatCtr.Add_GotFocus({ Paint-FocusBorder $this; $txt_Info.text = ($txt_repeatCtr.Tag) })
  $txt_repeatCtr.Add_LostFocus({ Paint-FocusBorder $this; $script:repeatCtr=$txt_repeatCtr.Text})
  $txt_repeatCtr.Add_MouseEnter({ Show-ToolTip $this })
  $txt_repeatCtr.Add_MouseLeave({ $obj_tt.Hide($form) })
  $txt_repeatCtr.Tag = "# of times VidMonHB should repeat execution (with the same options)."


  $lbl_Info                        = New-Object system.Windows.Forms.Label
  $lbl_Info.text                   = "Information"
  $lbl_Info.AutoSize               = $true
  $lbl_Info.width                  = 25
  $lbl_Info.height                 = 10
  $lbl_Info.location               = New-Object System.Drawing.Point(50,655)
  $lbl_Info.Font                   = $v12.font
  $lbl_Info.ForeColor              = $lblColor

  $txt_Info                        = New-Object System.Windows.Forms.RichTextBox
  $txt_Info.TabStop                = $false
  $txt_Info.AcceptsTab             = $true
  $txt_Info.AcceptsReturn          = $true
  $txt_Info.Multiline              = $true
  $txt_Info.width                  = 1237
  $txt_Info.height                 = 75
  $txt_Info.location               = New-Object System.Drawing.Point(160,637)
  $txt_Info.Font                   = $v12.font
  $txt_Info.ForeColor              = $blue
  $txt_Info.ReadOnly               = $true
  $txt_Info.ScrollBars             = 3 #ScrollBars.Vertical
  #$txt_Info.Add_GotFocus({  Paint-FocusBorder $this })
  #$txt_Info.Add_LostFocus({ Paint-FocusBorder $this })

  $lbl_propfileloc                 = New-Object system.Windows.Forms.Label
  $lbl_propfileloc.text            = "Cfg/Parms file"
  $lbl_propfileloc.AutoSize        = $true
  $lbl_propfileloc.width           = 25
  $lbl_propfileloc.height          = 10
  $lbl_propfileloc.BorderStyle =
  $lbl_propfileloc.location        = New-Object System.Drawing.Point(36,730)
  $lbl_propfileloc.Font            = $v12.font
  $lbl_propfileloc.ForeColor       = $lblColor

  $cbx_propfileloc                 = New-Object system.Windows.Forms.ComboBox
  $cbx_propfileloc.TabIndex        = 26
  $cbx_propfileloc.text            = $propfileloc
  $cbx_propfileloc.width           = 415
  $cbx_propfileloc.height          = 20
  $cbx_propfileloc.location        = New-Object System.Drawing.Point(160,730)
  $cbx_propfileloc.Font            = $ss12.font
  $cbx_propfileloc.ForeColor       = $blue
  $cbx_propfileloc.AutoCompleteMode= 3 #SuggestAppend
  $script:cbx_propfileloc_CurVal   = $cbx_propfileloc.Text
  $cbx_propfileloc.Items.Clear()
  $cbx_propfileloc.Text            = ""
  $items = Get-ChildItem "$currentEnv\*.ps-properties"
  $items | ForEach-Object {[void] $cbx_propfileloc.Items.Add($_.FullName)}
  $cbx_propfileloc.Text            = $script:cbx_propfileloc_CurVal
  $length = $items.FullName | Measure-Object -Maximum -Property Length | Select-Object maximum
  $cbx_propfileloc.DropDownWidth = 415 * $length.Maximum / 25
  $cbx_propfileloc.Add_GotFocus({  Paint-FocusBorder $this
                                   $txt_Info.text ="Select new configuration from drop down." 
                                   $script:cbx_propfileloc_CurVal   = $cbx_propfileloc.Text
                                   $cbx_propfileloc.Items.Clear()
                                   $cbx_propfileloc.Text            = ""
                                   $items = Get-ChildItem "$currentEnv\*.ps-properties" 
                                   $items | ForEach-Object {[void] $cbx_propfileloc.Items.Add($_.FullName)}
                                   $cbx_propfileloc.Text            = $script:cbx_propfileloc_CurVal
                                  })
  $cbx_propfileloc.Add_LostFocus({ Paint-FocusBorder $this
                                   $script:propfileloc=$cbx_propfileloc.text
                                   chkAllPaths
                                   $txt_info.backColor = $white
                                   $txt_info.foreColor = $blue
                                })
  $cbx_propfileloc.Add_SelectedIndexChanged({
    if ($cbx_propfileloc.SelectedItem -ne $script:cbx_propfileloc_CurVal) {
      # $a = new-object -comobject wscript.shell 
      # $intAnswer = $a.popup("Do you want to load new parameter values from " + $cbx_propfileloc.Text,"Replace Parms",4132) 
      # if ($intAnswer -in ("-1","6")) {
      $txt_Info.AppendText("`n")
      $txt_Info.AppendText("Loaded new config from " + $cbx_propfileloc.Text)
      $txt_Info.AppendText("`n")
      #Scroll to the end of the textbox
      $txt_Info.SelectionStart = $txt_Info.TextLength;
      $txt_Info.ScrollToCaret()
      loadConfig($cbx_propfileloc.Text)
      $script:cbx_propfileloc_CurVal          = $cbx_propfileloc.SelectedItem
      # }
      # else {
      #   $cbx_propfileloc.Text = $script:cbx_propfileloc_CurVal
      # }
      }
    })

  $btn_saveConfig                  = New-Object system.Windows.Forms.Button
  $btn_saveConfig.TabIndex         = 27
  $btn_saveConfig.BackColor        = $white
  $btn_saveConfig.text             = "Save Config/Parms"
  $btn_saveConfig.width            = 229
  $btn_saveConfig.height           = 30
  $btn_saveConfig.location         = New-Object System.Drawing.Point(590,730)
  $btn_saveConfig.Font             = $v14b.Font
  $btn_saveConfig.ForeColor        = $blue
  $btn_saveConfig.Add_Click({saveConfig})
  $btn_saveConfig.Add_GotFocus({  Paint-FocusBorder $this; $txt_Info.text ="Save configuration" })
  $btn_saveConfig.Add_LostFocus({ Paint-FocusBorder $this })

  $lbl_errorMsg                    = New-Object system.Windows.Forms.Label
  $lbl_errorMsg.text               = "Corrections`n    needed !"
  $lbl_errorMsg.AutoSize           = $true
  $lbl_errorMsg.width              = 25
  $lbl_errorMsg.height             = 10
  $lbl_errorMsg.BorderStyle        =
  $lbl_errorMsg.location           = New-Object System.Drawing.Point(975,723)
  $lbl_errorMsg.Font               = $v12b.font
  $lbl_errorMsg.ForeColor          = $red
  $lbl_errorMsg.BackColor          = $yellow
  $lbl_errorMsg.visible            = $false

  $btn_execute                     = New-Object system.Windows.Forms.Button
  $btn_execute.TabIndex            = 28
  $btn_execute.BackColor           = $white
  $btn_execute.text                = "Execute Script"
  $btn_execute.width               = 187
  $btn_execute.height              = 30
  $btn_execute.location            = New-Object System.Drawing.Point(1115,730)
  $btn_execute.Font                = $v14b.Font
  $btn_execute.ForeColor           = $blue
  $btn_execute.Add_GotFocus({  Paint-FocusBorder $this; $txt_Info.text ="Execute script with the parameters shown" })
  $btn_execute.Add_LostFocus({ Paint-FocusBorder $this })
  $btn_execute.Add_Click({$script:exit = $false; $form.Dispose()})

  $btn_exit                        = New-Object system.Windows.Forms.Button
  $btn_exit.TabIndex               = 29
  $btn_exit.BackColor              = $white
  $btn_exit.text                   = "Exit"
  $btn_exit.width                  = 70
  $btn_exit.height                 = 30
  $btn_exit.location               = New-Object System.Drawing.Point(1320,730)
  $btn_exit.Font                   = $v14b.Font
  $btn_exit.ForeColor              = $blue
  $btn_exit.DialogResult           = "Cancel"
  $btn_exit.Add_Click({$script:exit = $true;$form.Dispose()})
  $btn_exit.Add_GotFocus({  Paint-FocusBorder $this; $txt_Info.text ="Exit" })
  $btn_exit.Add_LostFocus({ Paint-FocusBorder $this })

  $gbx_convInfo.controls.AddRange(@($lbl_ParallelProcMax,$lbl_ParallelProcMsg,$lbl_limit,$lbl_vidtypes,$gbx_outSameAsIn,$gbx_movefiles,$gbx_delAfterConv,$txt_vidtypes,$txt_limit,$cbx_ParallelProcMax))
  $gbx_hbInfo.controls.AddRange(@($lbl_moviePreset,$lbl_tvPreset,$lbl_hbopts,$lbl_hbpreloc,$lbl_hbloc,$txt_hbloc,$txt_hbpreloc,$txt_hbopts,$cbx_tvPreset,$cbx_moviePreset))
  $gbx_location.controls.AddRange(@($lbl_in,$lbl_out,$lbl_out2,$lbl_logfilePath,$lbl_TVShowBasePath,$lbl_movieBasePath,$lbl_logfilePath,$txt_in,$txt_out,$txt_logfilePath,$txt_TVShowBasePath,$txt_movieBasePath))  
  $form.controls.AddRange(@($gbx_title,$gbx_location,$gbx_hbInfo,$gbx_convInfo,$gbx_postInfo,$btn_exit,$btn_execute,$btn_saveConfig,$lbl_Info,$txt_Info,$lbl_propfileloc,$cbx_propfileloc,$lbl_errorMsg))
  $gbx_title.controls.AddRange(@($lbl_heading,$lbl_version,$lbl_title,$pbx_image))
  $gbx_outSameAsIn.controls.AddRange(@($lbl_outSameAsIn,$cbx_outSameAsIn,$lbl_outSameAsIn2))
  $gbx_movefiles.controls.AddRange(@($lbl_movefiles,$lbl_movefiles2,$cbx_movefiles))
  $gbx_delAfterConv.controls.AddRange(@($cbx_delAfterConv,$lbl_delAfterConv))
  $gbx_postInfo.controls.AddRange(@($lbl_postExecCmd,$txt_postExecCmd,$lbl_postExecArgs,$txt_postExecArgs,$lbl_postNotify,$cbx_postNotify,$lbl_postLog,$cbx_postLog,$lbl_smtpServer,$txt_smtpServer,$lbl_smtpFromEmail,$txt_smtpFromEmail,$lbl_smtpToEmail,$txt_smtpToEmail,$gbx_postNotification,$lbl_repeatCtr,$txt_repeatCtr))

  function Paint-FocusBorder([System.Windows.Forms.Control]$control) {
      # get the parent control (usually the form itself)
      $parent = $control.Parent
      $parent.Refresh()
      if ($control.Focused) {
          $control.BackColor = $lightblue
          $pen = [System.Drawing.Pen]::new($red, 2)
        }
      else {
          $control.BackColor = $white
          $pen = [System.Drawing.Pen]::new($parent.BackColor, 2)
      }
      $rect = [System.Drawing.Rectangle]::new($control.Location, $control.Size)
      $rect.Inflate(1,1)
      $parent.CreateGraphics().DrawRectangle($pen, $rect)
  }

  function Show-ToolTip {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Windows.Forms.Control]$control,
        [string]$text = $null,
        [int]$duration = 10000
    )
    $script:txt_Info.AppendText($text)
    $script:txt_Info.AppendText("`n")
    #Scroll to the end of the textbox
    $script:txt_Info.SelectionStart = $txt_Info.TextLength;
    $script:txt_Info.ScrollToCaret()
    # Currently not fully functional. Will return to this in the future
    return
    if ([string]::IsNullOrWhiteSpace($text)) { $text = $control.Tag }
    $pos = [System.Drawing.Point]::new($control.Right, $control.Top)
    $obj_tt.Show($text, $control.Parent, $pos, $duration)
  }

  # Perform tasks the first time the form opens
  # call the Paint-FocusBorder when the form is first drawn
  $form.Add_Shown({
    Paint-FocusBorder $txt_in
    chkPath $txt_in
    chkPath $txt_out
    chkPath $txt_logfilePath
    chkPath $txt_hbpreloc
    chkPath $txt_hbloc
    chkPath $txt_TVShowBasePath
    chkPath $txt_movieBasePath
    chkPath $cbx_propfileloc
    chkPreset $cbx_tvPreset
    chkPreset $cbx_moviePreset
    if ($ParallelProcMax -gt 1) {
      $lbl_ParallelProcMsg.text = "Parallel Processing Enabled"} else {$lbl_ParallelProcMsg.text = ""}
    if ($cbx_outSameAsIn.Checked -eq $true) {$txt_out.Visible=$false; $lbl_out2.Visible=$true }
      else {$txt_out.Visible=$true; $lbl_out2.Visible=$false}
    if ($cbx_postNotify.SelectedItem -eq "None") {  
        $txt_smtpServer.Enabled          = $false
        $txt_smtpFromEmail.Enabled       = $false
        $txt_smtpToEmail.Enabled         = $false
      }
    })

  #Check if the preset exists within the presets.json file
  function chkPreset($object){
    $presetCheck = '"PresetName": "' + $object.text + '",'
    if ((Select-String -Path $script:hbpreloc -Pattern $presetCheck)) {
      $object.forecolor = $blue
      $object.backcolor = $white
      return $true
    }
    else {
      $object.forecolor = $red
      $object.backcolor = $yellow
      return $false
    }
  }

#Check if the path exists
function chkPath($object){
  if (Test-Path $object.text) {
    $object.forecolor = $blue
    $object.backcolor = $white
    return $true
  }
  else {
    $object.forecolor = $red
    $object.backcolor = $yellow
    return $false
  }
}

  #Check all paths
  function chkAllPaths {
    $validPath = $true
    if (-not (chkPath $txt_in)) {$validPath = $false}
    if (-not (chkPath $txt_out)) {$validPath = $false}
    if (-not (chkPath $txt_logfilePath)) {$validPath = $false}
    if (-not (chkPath $txt_hbpreloc)) {$validPath = $false}
    if (-not (chkPath $txt_hbloc)) {$validPath = $false}
    if (-not (chkPath $txt_TVShowBasePath)) {$validPath = $false}
    if (-not (chkPath $txt_movieBasePath)) {$validPath = $false}
    if (-not (chkPath $cbx_propfileloc)) {$validPath = $false}
    if ($validPath -eq $true) {$lbl_errorMsg.Visible = $false}
    else {$lbl_errorMsg.Visible = $true}
  }

  #Write your logic code here
  [void]$form.ShowDialog()
  #[System.Windows.Forms.Application]::Run($form)
  # clean-up
  $obj_tt.Dispose()
  $form.Dispose()
} #displayForm

#Replace the parameters with the values from a different parameter file
#Note - This can only be called from the displayForm function
function loadConfig($parmFile) {
  try {
    if (Test-Path $parmFile) {
      # Convert properties file to hashtable that can be searched by value
      $properties = Get-Content $parmFile -Raw  | ConvertFrom-StringData
      # Copy value to new hashtable, but expand env vars first
      $expanded = @{}
      foreach($entry in $properties.GetEnumerator()){
          $expanded[$entry.Key] = [Environment]::ExpandEnvironmentVariables($entry.Value)
      }  
      $script:vidTypes        = $expanded["vidTypes"]
      $txt_vidTypes.text      = $script:vidTypes
      $script:in              = $expanded["in"]
      $txt_in.text            = $script:in
      $script:out             = $expanded["out"]
      $txt_out.text           = $script:out
      [Boolean]$script:outSameAsIn = [System.Convert]::ToBoolean($expanded["outSameAsIn"])
      $cbx_outSameAsIn.Checked= $script:outSameAsIn
      $script:delAfterConv    = $expanded["delAfterConv"]
      $cbx_delAfterConv.text  = $script:delAfterConv
      $script:hbloc           = $expanded["hbloc"]
      $txt_hbloc.text         = $script:hbloc
      $script:hbpreloc        = $expanded["hbpreloc"]
      $txt_hbpreloc.text      = $script:hbpreloc
      $script:tvPreset        = $expanded["tvPreset"]
      $cbx_tvPreset.text      = $script:tvPreset
      $script:moviePreset     = $expanded["moviePreset"]
      $cbx_moviePreset.text   = $script:moviePreset
      $script:hbopts          = $expanded["hbopts"]
      $txt_hbopts.text        = $script:hbopts
      [Boolean]$script:movefiles = [System.Convert]::ToBoolean($expanded["movefiles"])
      $cbx_movefiles.Checked  = $script:movefiles
      $script:logfilePath     = $expanded["logfilePath"]
      $txt_logfilePath.text   = $script:logfilePath
      $script:TVShowBasePath  = $expanded["TVShowBasePath"]
      $txt_TVShowBasePath.text= $script:TVShowBasePath
      $script:movieBasePath   = $expanded["movieBasePath"]
      $txt_movieBasePath.text = $script:movieBasePath
      [Int]$script:ParallelProcMax = $expanded["ParallelProcMax"]
      $cbx_ParallelProcMax.text = $script:ParallelProcMax
      $script:limit           = $expanded["limit"]
      $txt_limit.text         = $script:limit
      $script:postExecCmd     = $expanded["postExecCmd"]
      $txt_postExecCmd.text   = $script:postExecCmd
      $script:postExecArgs    = $expanded["postExecArgs"]
      $txt_postExecArgs.text  = $script:postExecArgs
      $script:postNotify      = $expanded["postNotify"]
      $cbx_postNotify.text    = $script:postNotify
      $script:smtpServer      = $expanded["smtpServer"]
      $txt_smtpServer.text    = $script:smtpServer
      $script:smtpFromEmail   = $expanded["smtpFromEmail"]
      $txt_smtpFromEmail.text = $script:smtpFromEmail
      $script:smtpToEmail     = $expanded["smtpToEmail"]
      $txt_smtpToEmail.text   = $script:smtpToEmail
      if ($script:ParallelProcMax -gt 1) {
        $lbl_ParallelProcMsg.text = "Parallel Processing Enabled"} else {$lbl_ParallelProcMsg.text = ""}
      if ($cbx_outSameAsIn.Checked -eq $true) {$txt_out.Visible=$false; $lbl_out2.Visible=$true }
        else {$txt_out.Visible=$true; $lbl_out2.Visible=$false} }
      $script:postLog         = $expanded["postLog"]
      $cbx_postLog.text       = $script:postLog
      chkAllPaths
      chkPreset $cbx_tvPreset
      chkPreset $cbx_moviePreset
    }
  catch {
    $errorMsg = $_.Exception.Message
    #Invalid Pattern error comes from videoFiles information. Ignore
    $txt_Info.AppendText("Error trying to read parameter file in this folder.  $parmFile - $errorMsg")
    $txt_Info.AppendText("`n")
    #Scroll to the end of the textbox
    $txt_Info.SelectionStart = $txt_Info.TextLength;
    $txt_Info.ScrollToCaret()
    return
  }
} #loadConfig

#Delete, Recycle or maintain the original video file 
function delFile($fname) {
  if ($delAfterConv -eq "Recycle") {
    writeLog ("Recycled file " + $fname)
    [FileIO.FileSystem]::DeleteFile($fname, 'OnlyErrorDialogs', 'SendToRecycleBin')
  }
  else {
    writeLog ("Deleted file " + $fname)
    Remove-Item $fname -Force -ErrorAction SilentlyContinue
  }
}

#Force recycle the file (used for log files)
function recycleFile($fname) {
    writeLog ("Recycled log file " + $fname)
    [FileIO.FileSystem]::DeleteFile($fname, 'OnlyErrorDialogs', 'SendToRecycleBin')
}

# Display conversion history information (daily, monthly, yearly)
function displayHistory() {
  $HistoryLogFile = Join-Path -Path $logFilePath -ChildPath "VidMonHB_History.csv"
  $historyCSV = Import-Csv $HistoryLogFile
  $yyyy = (get-date -Format yyyy)
  $mm = (get-date -Format MM)
  $dd = (get-date -Format dd)
  $historyDailyBegSize = 0
  $historyDailyEndSize = 0
  $historyDailyPct = 0
  $historyDailyFileCount = 0
  $historyDailyPercentage = 0
  $historyMonthlyBegSize = 0
  $historyMonthlyEndSize = 0
  $historyMonthlyPct = 0
  $historyMonthlyFileCount = 0
  $historyYearlyBegSize = 0
  $historyYearlyEndSize = 0
  $historyYearlyPct = 0
  $historyYearlyFileCount = 0
  foreach ($historyItem in $historyCSV) {
    # Daily
    if (($historyItem.yyyy -eq $yyyy) -and 
        ([int]$historyItem.mm -eq [int]$mm ) -and 
        ([int]$historyItem.dd -eq [int]$dd ))
    {
      $historyDailyBegSize += $historyItem.BegSize
      $historyDailyEndSize += $historyItem.EndSize
      $historyDailyPct = [string]([math]::Round(100-($historyDailyEndSize / $historyDailyBegSize)*100,0)) + "%"
      $historyDailyFileCount += $historyItem.FileCount
    }
    # Monthly
    if (($historyItem.yyyy -eq $yyyy) -and 
        ([int]$historyItem.mm -eq [int]$mm ))
    {
      $historyMonthlyBegSize += $historyItem.BegSize
      $historyMonthlyEndSize += $historyItem.EndSize
      $historyMonthlyPct = [string]([math]::Round(100-($historyMonthlyEndSize / $historyMonthlyBegSize)*100,0)) + "%"
      $historyMonthlyFileCount += $historyItem.FileCount
    }
    # Yearly
    if (($historyItem.yyyy -eq $yyyy))
    {
      $historyYearlyBegSize += $historyItem.BegSize
      $historyYearlyEndSize += $historyItem.EndSize
      $historyYearlyPct = [string]([math]::Round(100-($historyYearlyEndSize / $historyYearlyBegSize)*100,0)) + "%"
      $historyYearlyFileCount += $historyItem.FileCount
    }
  }

  Write-Color -LinesBefore 10  "     Yearly Totals    " -Color Blue -BackGroundColor Gray
  Write-Color "Beginning Size: ",('{0:n0}' -f $historyYearlyBegSize)," GB" -Color Yellow, Yellow, Yellow -BackGroundColor DarkCyan, DarkCyan, DarkCyan
  Write-Color "   Ending Size: ",('{0:n0}' -f $historyYearlyEndSize)," GB" -Color Yellow, Yellow, Yellow -BackGroundColor DarkCyan, DarkCyan, DarkCyan
  if ($historyYearlyPct -ge 0) {
    Write-Color "  Disk Savings: ", $historyYearlyPct -Color Cyan, Cyan -BackGroundColor DarkCyan, DarkCyan }
  else {        
    Write-Color "     Disk Loss: ", $historyYearlyPct -Color Cyan, Cyan -BackGroundColor DarkCyan, DarkCyan }
  Write-Color "    File Count: ",('{0:n0}' -f $historyYearlyFileCount) -Color Yellow, Yellow -BackGroundColor DarkCyan, DarkCyan

  Write-Color -LinesBefore 2   "    Monthly Totals    " -Color Blue -BackGroundColor Gray
  Write-Color "Beginning Size: ",('{0:n0}' -f $historyMonthlyBegSize)," GB" -Color Yellow, Yellow, Yellow -BackGroundColor DarkCyan, DarkCyan, DarkCyan
  Write-Color "   Ending Size: ",('{0:n0}' -f $historyMonthlyEndSize)," GB" -Color Yellow, Yellow, Yellow -BackGroundColor DarkCyan, DarkCyan, DarkCyan
  if ($historyMonthlyPct -ge 0) {
    Write-Color "  Disk Savings: ", $historyMonthlyPct -Color Cyan, Cyan -BackGroundColor DarkCyan, DarkCyan }
  else {        
    Write-Color "     Disk Loss: ", $historyMonthlyPct -Color Cyan, Cyan -BackGroundColor DarkCyan, DarkCyan }
  Write-Color "    File Count: ",('{0:n0}' -f $historyMonthlyFileCount) -Color Yellow, Yellow -BackGroundColor DarkCyan, DarkCyan

  Write-Color -LinesBefore 2   "      Daily Totals    " -Color Blue -BackGroundColor Gray
  Write-Color "Beginning Size: ",('{0:n0}' -f $historyDailyBegSize)," GB" -Color Yellow, Yellow, Yellow -BackGroundColor DarkCyan, DarkCyan, DarkCyan
  Write-Color "   Ending Size: ",('{0:n0}' -f $historyDailyEndSize)," GB" -Color Yellow, Yellow, Yellow -BackGroundColor DarkCyan, DarkCyan, DarkCyan
  if ($historyDailyPct -ge 0) {
    Write-Color "  Disk Savings: ", $historyDailyPct -Color Cyan, Cyan -BackGroundColor DarkCyan, DarkCyan }
  else {        
    Write-Color "     Disk Loss: ", $historyDailyPct -Color Cyan, Cyan -BackGroundColor DarkCyan, DarkCyan }
  Write-Color "    File Count: ",('{0:n0}' -f $historyDailyFileCount) -Color Yellow, Yellow -BackGroundColor DarkCyan, DarkCyan
}

#-----------------------------------------[Start Execution]------------------------------------------

Clear-Host

if ($Host.Name -eq "Visual Studio Code Host") {$bgColor="Black"}
else {$bgColor = [System.Console]::BackgroundColor}

#Check if the user wants to recycle instead of deleting completed files
$recycleAvailble = $true
if ( -not (Get-Module -ListAvailable -Name Recycle)) {
  $recycleAvailble = $false
  writeLog "Warning - The Powershell Recycle module has not installed." -logType "S" -logSeverity "E"
  writeLog "Install-Module -Name Recycle -RequiredVersion 1.0.2 -Scope CurrentUser -Force" -logType "S" -logSeverity "E"
  $a = new-object -comobject wscript.shell 
  $intAnswer = $a.popup("Warning - The Powershell Recycle module has not installed.`n`nPlease run the following command:`n`n" + 
                        "Install-Module -Name Recycle -RequiredVersion 1.0.2 -Scope CurrentUser -Force",15,"Recycle Module Missing",4096) 
  #exit
}

#Check if the user wants to recycle instead of deleting completed files
$PSWriteColor = $true
if ( -not (Get-Module -ListAvailable -Name PSWriteColor)) {
  $PSWriteColor = $false
  writeLog "Warning - The Powershell PSWriteColor module has not installed." -logType "S" -logSeverity "E"
  writeLog "Install-Module -Name PSWriteColor -Scope CurrentUser -Force" -logType "S" -logSeverity "E"
  $a = new-object -comobject wscript.shell 
  $intAnswer = $a.popup("Warning - The Powershell PSWriteColor module has not installed.`n`nPlease run the following command:`n`n" + 
                        "Install-Module -Name PSWriteColor -Scope CurrentUser -Force",15,"PSWriteColor Module Missing",4096) 
  #exit
}

#Check if the prior session did not complete and if we want to resume
if (Test-Path $resumeFile) {
  [console]::beep(1000,200)
  [console]::beep(1000,200)
  writeLog ("Warning - Previous session did not complete.") -logType "S" -logSeverity "E"
  $a = new-object -comobject wscript.shell 
  $intAnswer = $a.popup("Warning - Previous session did not complete.`nDo you want to resume the "+
                        "previous session (Y/N)?`n`nTimeout will default to Yes in 15 seconds",15,"Resume Question",4132) 
  if ($intAnswer -in ("-1","6")) {
    $resume = $true
  }
  else {Remove-Item $resumeFile}
}

if (($resume) -and (checkFileLocked($resumeFile))) {
  writeLog ("Warning - resume file $resumeFile is locked. " +
             "This session will not be able to resume if an error occurs") -logType "S" -logSeverity "E"
  do {
    $response = Read-Host -Prompt "Do you want to continue anyway (Y/N)"
    $response = $response.ToUpper()
    if ($response -eq "N") {
      writeLog "Exiting" -logType "S" 
      exit
    }
  } until ($response -eq "Y")
}

#Read the properties file and replace the parameters with specified values
#If we are resuming from a previous session, read the parameters from the resume file
$parmFile = $propfileloc
if ($resume) {
  #Resuming - Parameters will come from the resume file
  $parmFile = $resumeFile + ".temp"
  #Pull out just the parameters, not the videofile resume information
  $lineNum = Select-String -Pattern "videofile" -Path $resumeFile -list -SimpleMatch | select-object -First 1
  Get-Content $resumeFile -TotalCount ($lineNum.LineNumber-1) | Set-Content $parmFile
}

try {
  if (Test-Path $parmFile) {
    # Convert properties file to hashtable that can be searched by value
    $properties = Get-Content $parmFile -Raw  | ConvertFrom-StringData
    # Copy value to new hashtable, but expand env vars first
    $expanded = @{}
    foreach($entry in $properties.GetEnumerator()){
        $expanded[$entry.Key] = [Environment]::ExpandEnvironmentVariables($entry.Value)
    }  
    $vidTypes        = $expanded["vidTypes"]
    $in              = $expanded["in"]
    $out             = $expanded["out"]
    [Boolean]$outSameAsIn = [System.Convert]::ToBoolean($expanded["outSameAsIn"])
    $delAfterConv    = $expanded["delAfterConv"]
    $hbloc           = $expanded["hbloc"]
    $hbpreloc        = $expanded["hbpreloc"]
    $tvPreset        = $expanded["tvPreset"]
    $moviePreset     = $expanded["moviePreset"]
    $hbopts          = $expanded["hbopts"]
    [Boolean]$movefiles = [System.Convert]::ToBoolean($expanded["movefiles"])
    $logfilePath     = $expanded["logfilePath"]
    $TVShowBasePath  = $expanded["TVShowBasePath"]
    $movieBasePath   = $expanded["movieBasePath"]
    $ParallelProcMax = $expanded["ParallelProcMax"]
    $limit           = $expanded["limit"]
    $postExecCmd     = $expanded["postExecCmd"]
    $postExecArgs    = $expanded["postExecArgs"]
    $postNotify      = $expanded["postNotify"]
    $smtpServer      = $expanded["smtpServer"]
    $smtpFromEmail   = $expanded["smtpFromEmail"]
    $smtpToEmail     = $expanded["smtpToEmail"]
    $postLog         = $expanded["postLog"]
  }
}
catch {
  $errorMsg = $_.Exception.Message
  #Invalid Pattern error comes from videoFiles information. Ignore
  writeLog "Error trying to read parameter file in this folder.  $parmFile" -logType "S" -logSeverity "E"
  writeLog ("$errorMsg") -logType "S" -logSeverity "E"
  writeLog "Exiting" -logType "S" -logSeverity "E"
  return
}

#Remove the temp file created during resume mode
if ($resume) {Remove-Item $parmFile -Force -ErrorAction SilentlyContinue}

if ($entryType -ne "winForm") {displayParms("S")}

#http://msdn.microsoft.com/en-us/library/x83z1d9f(v=vs.84).aspx
# Check if repeating.
if ($repeatCtr -gt 0) {$repeatCtr = $repeatCtr-1}
else {
  $a = new-object -comobject wscript.shell 
  $intAnswer = $a.popup("Press the Enter key to edit parameters.`n`nTimeout in 5 seconds",5,"Edit Parameters",4096) 

  if ($intAnswer -gt -1) {
    if ($entryType -eq 'winForm') {
        displayForm
        if ($exit -eq $true) {writeLog "Exiting" -logType "S"; exit}
        if ($repeatCtr -gt 0) {$repeatCtr = $repeatCtr-1}
      }
    else {
      if ($intAnswer -gt -1) {
        $option = "loop"
        while ($null -ne $option -and $option -notin ("88", "99")) {
          displayParms("S")
          $option = Read-Host -Prompt "Enter parameter change"
          if ([string]::IsNullOrWhiteSpace($option)) { $option = "99" }
          switch ($option) {
            "1" {
              $result = (Read-Host -Prompt "Video types [$vidTypes]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $vidTypes = $result }
            }
            "2" {
              $result = (Read-Host -Prompt "Input location [$in]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $in = $result }
              if ( -not $in.Endswith("\")) { $in += "\" }
            }
            "3" {
              $result = (Read-Host -Prompt "Output location [$out]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $out = $result }
              if ( -not $out.Endswith("\")) { $out += "\" }
            }
            "4" {
              $result = (Read-Host -Prompt "Same As Input Override [$outSameAsIn]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) {
                if ($result.ToUpper() -eq "TRUE") { $outSameAsIn = $true }
                if ($result.ToUpper() -eq "FALSE") { $outSameAsIn = $false }
              }
            }
            "5" {
              $result = (Read-Host -Prompt "Delete original [$delAfterConv]").Replace("`"", "").ToUpper()
              if ( -not [string]::IsNullOrWhiteSpace($result) -and ($result -in ("Maintain", "Delete", "Recycle"))) {
                $delAfterConv = $result.Trim()
              }
            }
          "6" {
              $result = (Read-Host -Prompt "HandBrake file location [$hbloc]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $hbloc = $result }
            }
            "7" {
              $result = (Read-Host -Prompt "HB preset file location [$hbpreloc]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $hbpreloc = $result }
            }
            "8" {
              $result = (Read-Host -Prompt "HandBrake presets [$tvPreset]").Replace("`"", "")
              $presetCheck = '"PresetName": "' + $result + '",'
              if ( -not [string]::IsNullOrWhiteSpace($result)) {
                if (Select-String -Path $hbpreloc -Pattern $presetCheck) {
                  $tvPreset = $result
                }
                else {
                  [console]::beep(1000, 200)
                  [console]::beep(1000, 200)
                  writeLog ('Error - A preset named "' + $result + '" could not be found in preset file ' + $hbpreloc) -logType "S" -logSeverity "E"
                  Start-Sleep 4
                }
              }
            }
            "9" {
              $result = (Read-Host -Prompt "HandBrake presets [$moviePreset]").Replace("`"", "")
              $presetCheck = '"PresetName": "' + $result + '",'
              if ( -not [string]::IsNullOrWhiteSpace($result)) {
                if (Select-String -Path $hbpreloc -Pattern $presetCheck) {
                  $moviePreset = $result
                }
                else {
                  [console]::beep(1000, 200)
                  [console]::beep(1000, 200)
                  writeLog ('Error - A preset named "' + $result + '" could not be found in preset file ' + $hbpreloc) -logType "S" -logSeverity "E"
                  Start-Sleep 4
                }
              }
            }
            "10" {
              $result = (Read-Host -Prompt "HandBrake options [$hbopts]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $hbopts = $result }
            }
            "11" {
              $result = (Read-Host -Prompt "Property location [$propfileloc]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $propfileloc = $result }
            }
            "12" {
              $result = (Read-Host -Prompt "Move files option [$movefiles]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) {
                if ($result.ToUpper() -eq "TRUE") { $movefiles = $true }
                if ($result.ToUpper() -eq "FALSE") { $movefiles = $false }
              }
            }
            "13" {
              $result = (Read-Host -Prompt "Log file path [$logfilePath]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $logfilePath = $result }
            }
            "14" {
              $result = (Read-Host -Prompt "TV Show base path [$TVShowBasePath]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $TVShowBasePath = $result }
              if ( -not $TVShowBasePath.Endswith("\")) { $TVShowBasePath += "\" }
            }
            "15" {
              $result = (Read-Host -Prompt "Movie base path [$movieBasePath]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $movieBasePath = $result }
              if ( -not $movieBasePath.Endswith("\")) { $movieBasePath += "\" }
            }
            "16" {
              $result = (Read-Host -Prompt "Parallel processing max [$ParallelProcMax]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result) -and ([int]$result -in 1..10)) {
                $ParallelProcMax = $result
              }
            }
            "17" {
              $result = (Read-Host -Prompt "# of files to process [$limit]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result) -and ([int]$result -in 0..999)) {
                $limit = $result
              }
            }
            "18" {
              $result = (Read-Host -Prompt "Post exec cmd [$postExecCmd]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $postExecCmd = $result }
            }
            "19" {
              $result = (Read-Host -Prompt "Post exec arguments [$postExecArgs]").Replace("`"", "")
              if ( -not [string]::IsNullOrWhiteSpace($result)) { $postExecArgs = $result }
            }
            "20" {
              $result = (Read-Host -Prompt "Post notification [$postNotify]").Replace("`"", "").ToUpper()
              if ( -not [string]::IsNullOrWhiteSpace($result) -and ($result -in ("None", "All", "Error"))) {
                $postNotify = $result.Trim()
              }
            }
            "77" {
              #save config changes back to the vidMonHB.ps-properties file
              saveConfig
            }
            "88" {
              writeLog "Exiting" -logType "S"
              if ( -not ($resume)) { Remove-Item $resumeFile -ErrorAction SilentlyContinue }
              exit
            }
            Default { $null }
          } 
        } # while
      } # if ($intAnswer -gt -1)
    }
  }
}
#Populate the resume parameters which will be written to the resumeVidMonHB.txt file.
#The videoFiles key will be updated as each file is completed
$resumeParms['vidTypes']=$vidTypes
$resumeParms['in']=($in).Replace("\","\\")
$resumeParms['out']=($out).Replace("\","\\")
$resumeParms['outSameAsIn']=$outSameAsIn
$resumeParms['delAfterConv']=$delAfterConv
$resumeParms['hbloc']=($hbloc).Replace("\","\\")
$resumeParms['hbpreloc']=($hbpreloc).Replace("\","\\")
$resumeParms['tvPreset']=$tvPreset
$resumeParms['moviePreset']=$moviePreset
$resumeParms['hbopts']=$hbopts
$resumeParms['movefiles']=$movefiles
$resumeParms['logfilePath']=($logfilePath).Replace("\","\\")
$resumeParms['TVShowBasePath']=($TVShowBasePath).Replace("\","\\")
$resumeParms['movieBasePath']=($movieBasePath).Replace("\","\\")
$resumeParms['ParallelProcMax']=$ParallelProcMax
$resumeParms['limit']=$limit
$resumeParms['postExecCmd']=($postExecCmd).Replace("\","\\")
$resumeParms['postExecArgs']=$postExecArgs
$resumeParms['postNotify']=$postNotify
$resumeParms['smtpServer']=$smtpServer
$resumeParms['smtpFromEmail']=$smtpFromEmail
$resumeParms['smtpToEmail']=$smtpToEmail
$resumeParms['postLog']=$postLog
if ($resume) {
  writeLog ("Warning - Previous run did not complete. Resuming") -logSeverity "E"
}
else {
  $resumeParms.GetEnumerator() | ForEach-Object {"{0}={1}" -f $_.Name,$_.Value} | Set-Content $resumeFile
}
 

#Check if the Output folders exist.  If not, create them.
if ( -not ($outSameAsIn)) {
  mkdir $out -ErrorAction SilentlyContinue
}
mkdir $logFilePath -ErrorAction SilentlyContinue

#Summary log file contains log data just from this script (does not include HandBrake)
$logName = "VidMonHB_" + $timestamp + "_Summary.txt"
$sumLogFile = Join-Path -Path $logFilePath -ChildPath $LogName
Remove-Item $sumLogFile -ErrorAction SilentlyContinue

if ($ParallelProcMax -gt 10) {$ParallelProcMax = 10} # 10 is the max allowed
$sleepAmt = 30 #sleep for x number of seconds

#Make sure $in ends with *
#if ($in.Substring($in.Length-1) -ne "*") {$in += "*"}

#Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion
#Clear-Host #`
writeLog "`n`n`t`t`tVidMonHB - Version $version"
writeLog ("-").PadRight(80,"-")
writeLog "Parameter Settings"
writeLog "Video types             : $vidTypes"
writeLog "Input location          : $in"
if ( -not ($outSameAsIn)) {
writeLog "Output location         : $out" }
else {
writeLog "Output location         : Override - Output will be written to the same folder where the input was found" }

writeLog "Same As Input Override  : $outSameAsIn"
writeLog "Delete original         : $delAfterConv"
writeLog "HandBrake file location : $hbloc"
writeLog "HB preset file location : $hbpreloc"
writeLog "HandBrake TV preset     : $tvPreset"
writeLog "HandBrake movie preset  : $moviePreset"
writeLog "HandBrake options       : $hbopts"
writeLog "Property location       : $propfileloc" 
writeLog "Move files option       : $movefiles" 
writeLog "Log file path           : $logfilePath" 
writeLog "TV Show base path       : $TVShowBasePath" 
writeLog "Movie base path         : $movieBasePath" 
switch ($ParallelProcMax) { {$_ -lt 2} {$ppmsg="Single Threaded Mode"  } Default {$ppmsg="Parallel Processing Mode Enabled"} }
writeLog "Parallel processing max : $ParallelProcMax - $ppmsg"
writeLog "# of files to process   : $limit"
writeLog "Post exec cmd           : $postExecCmd"
writeLog "Post exec arguments     : $postExecArgs"
writeLog "Post notification       : $postNotify"
writeLog "SMTP Server             : $smtpServer"
writeLog "From Email              : $smtpFromEmail"
writeLog "To Email                : $smtpToEmail"
writeLog "Post open log           : $postLog"

#The taglib-sharp.dll module is used to clear the title name metadata
$ClearMetaFlag = $false
if (Test-Path ".\taglib-sharp.dll") {
  Import-Module ".\taglib-sharp.dll"
  $ClearMetaFlag = $true
  writeLog "Taglib found            : Title metadata will be cleared"
} else { writeLog "Taglib not found        : Title metadata will NOT be cleared" }

if (Test-Path $propfileloc) {
  writeLog ("Note - Parameters Settings were pulled from " + $propfileloc)}
writeLog ("-").PadRight(80,"-")

#Run Pre-Execution verification checks

#Check to ensure input path exists
if ( -not (Test-Path "$in" )) {
  writeLog ("Input file path $in not found") -logSeverity "E"
  writeLog ("`nExiting.") -logSeverity "E"
  return
}

#Check to ensure output path exists
if ( -not (Test-Path "$out" )) {
  writeLog ("Output file path $out not found") -logSeverity "E"
  writeLog ("`nExiting.") -logSeverity "E"
  return
}

#Check to ensure HandBrakeCLI.exe is installed
if ( -not (Test-Path "$hbloc" )) {
  writeLog ("HandBrakeCLI.exe not found at $hbloc") -logSeverity "E"
  WriteLog ("Program can be downloaded from https://handbrake.fr/downloads2.php") -logSeverity "E"
  writeLog ("`nExiting.") -logSeverity "E"
  return
}

#Check to ensure presets file exists
if ( -not (Test-Path "$hbpreloc" )) {
  writeLog ("Presets file not found at $hbpreloc") -logSeverity "E"
  writeLog ("`nExiting.") -logSeverity "E"
  return
}

#Check to ensure presets file exists
if ($postExecCmd) {
  if ( -not (Test-Path "$postExecCmd" )) {
    writeLog ("Presets file not found at $postExecCmd") -logSeverity "E"
    writeLog ("`nExiting.") -logSeverity "E"
    return
  }
}

if ($movefiles) {
  #Check to ensure TVShowBasePath file directory exists
  if ( -not (Test-Path $TVShowBasePath )) { 
    writeLog ("TV Show base path not found at $TVShowBasePath") -logSeverity "E"
    writeLog ("`nExiting.") -logSeverity "E"
    return
  }

  #Check to ensure TVShowBasePath file directory exists
  if ( -not (Test-Path $movieBasePath )) { 
    writeLog ("TV Show base path not found at $movieBasePath") -logSeverity "E"
    writeLog ("`nExiting.") -logSeverity "E"
    return
  }
}

#Check if TV preset exists
$presetCheck = '"PresetName": "' + $tvPreset + '",'
if ( -not (Select-String -Path $hbpreloc -Pattern $presetCheck)) {
  writeLog ('Error - A preset named "' + $tvPreset + '" could not be found in preset file ' + $hbpreloc) -logType "S" -logSeverity "E"
  return
}

#Check if Movie preset exists
$presetCheck = '"PresetName": "' + $moviePreset + '",'
if ( -not (Select-String -Path $hbpreloc -Pattern $moviePreset)) {
  writeLog ('Error - A preset named "' + $moviePreset + '" could not be found in preset file ' + $hbpreloc) -logType "S" -logSeverity "E"
  return
}

#Split the parameter and then put it back together again with wildcards
$vidTypesSplit = ($vidTypes.Replace("`"","").Split(",")).trim()
$incFiles=@()
forEach ($vidType in $vidTypesSplit) 
    {if ($vidType -ne "") {$incFiles += "*.$vidtype"}}

#Find the matching files
if ($resume) {
  Get-Content $resumeFile | ForEach-Object {
    if ($_.Split("=")[0] -eq "Unprocessed videofile") {
      $videoFiles += [System.IO.FileInfo]$_.Split("=")[1]
    }
  }
}
else {
  Clear-Host
  Write-Color "`nSearching for files.  Hit Ctrl-C to break" -BackGroundColor DarkCyan -Color Yellow
  $videoFiles = Get-ChildItem -path $in -Recurse -include $incFiles | where Length -gt $minSize | Select-Object -First $limit  
}

$fileCount = ($videofiles | Measure-Object).Count
if ($fileCount -eq 0)
  {
    writeLog  "`nNo files found in $in `nExiting.`n`n"
    Remove-Item $resumeFile -ErrorAction SilentlyContinue
    # if repeatMonitor is set to true, just sit and loop until files are found
    while ($true -eq $repeatMonitor) {
      Clear-Host
      $date = (get-date).ToString("MM/dd/yyyy hh:mm tt")
      Write-Color "`n$date - VidMonHB Monitor mode - Searching for files.  Hit Ctrl-C to break" -BackGroundColor DarkCyan -Color Yellow
      $videoFiles = Get-ChildItem -path $in -Recurse -include $incFiles | where Length -gt $minSize
      $fileCount = ($videofiles | Measure-Object).Count
      if ($fileCount -gt 0) {Invoke-Expression -Command ($PSCommandPath + ' -repeatMonitor $true') ; return }
      Clear-Host
      displayHistory
      $i=60; 
      do {$i--
          $date = (get-date).ToString("MM/dd/yyyy hh:mm tt")
          Write-Progress -Activity "$date - VidMonHB Monitor mode - Will look for new files to process in 1 minute.  Hit Ctrl-C to break" -PercentComplete ($i/60*100) -Status "Seconds remaining $i"
          Start-Sleep -Seconds 1
      } while ($i -ne 0)
      Write-Progress -Activity "$date - VidMonHB Monitor mode - Will look for new files to process in 1 minute.  Hit Ctrl-C to break" -Completed
      Clear-Host
    }
    Return
  }

if ($resume) {
  writeLog "`nResume found following $fileCount file(s) to process:"
}
else {
  writeLog "`nSearch found the following $fileCount file(s) to process:"
}

$padSize = ([string]$fileCount).Length

$i=0
foreach ($file in $videoFiles) {
  $i++; 
  $fileAttrib=$null
  if ((($file).Attributes -band $readonly -eq $readonly)) {$fileAttrib = " - ReadOnly"}
  writeLog (([string]$i).PadLeft($padSize,'0') + " - " + """$file""" + $fileAttrib)
  if ( -not ($resume)) {
    ("Unprocessed videofile=" + $file.fullName) | Add-Content $resumeFile
  }
}
writeLog ""

$estBegSize=0
#Create an ETA based on the number of files (single threaded only)
if ($ParallelProcMax -le 1) {
  writeLog "`nNow Processing File(s) - Please stand by"
  $estBegSize = [math]::Round(($videoFiles | Measure-Object -Sum Length).Sum / 1GB,3)
  $estCompMins = [math]::Round($estBegSize*6.8855)
  $estCompTime = ((get-Date).AddMinutes($estCompMins)).ToString('MM/dd/yyyy hh:mm:ss tt')
  writeLog ("`n$fileCount files to process. Estimated # of mins to complete $estBegSize GB is " +
            "$estCompMins minutes.  ETA $estCompTime`n" )
}

if ( -not $in.Endswith("\")) { $in += "\" }
if ( -not $out.Endswith("\")) { $out += "\" }
if ( -not $movieBasePath.Endswith("\")) { $movieBasePath += "\" }
if ( -not $TVShowBasePath.Endswith("\")) { $TVShowBasePath += "\" }

### TESTING
#return  #TESTING
### TESTING

#Main processing logic. Start looping through each of the files and run them through HandBrake.
$i=0
foreach ($file in $videoFiles) {
  # If single threaded, alternate the color schemes for START and END msgs
  if ($ParallelProcMax -le 1) { if ($i % 2 -eq 0) {$msgBGcolor = "DarkCyan"} else {$msgBGcolor = "DarkGray"} }
  writeLog "****START****START****START****START****START****START****START****START****START****START****" -logBGcolor $msgBGcolor
  $i++
  $countMsg = (([string]$i).PadLeft($padSize,'0') + " of " + ([string]$fileCount).PadLeft($padSize,'0'))
  writeLog ("Now processing " + $countMsg)
  writeLog ("File Name : """ + $file + """")
  $begSize = [math]::Round(($file | Measure-Object -Sum Length).Sum / 1GB,3)  
  writeLog ("Start Time: " + ($((Get-Date).ToString()))) 
  writeLog ("Start size: " + $begSize + " GB") 
  clearTitleMeta($file.fullName)
  if ($outSameAsIn) {
    $folder = Split-Path $file -Parent
    if ( -not $folder.Endswith("\")) { $folder += "\" }
#    $newFileName = $folder + "\" + $file.baseName + ".mp4"
    $newFileName = $folder + $file.baseName + ".mp4"
  }
  else {
#    $newFileName = $out + "\" + $file.baseName + ".mp4"
    $newFileName = $out + $file.baseName + ".mp4"
  }
  
  #This will determine which preset file to use (defaults to Movie)
  if (checkIfTVfile($file)) {
    if ($tvPreset) {$tvPresetname = '--preset ' + '"' + $tvPreset + '"' } #Only pass tvPreset parameter if populated
    $cmdArgs = "-i `"$file`" -t 1 -o `"$newFileName`" --preset-import-file `"$hbpreloc`" $tvPresetName `"$hbopts`""
    writeLog "HB Command: $hbloc -ArgumentList $cmdArgs`n"
  }
  else {
    if ($moviePreset) {$moviePresetName = '--preset ' + '"' + $moviePreset + '"' } #Only pass moviePreset parameter if populated
    $cmdArgs = "-i `"$file`" -t 1 -o `"$newFileName`" --preset-import-file `"$hbpreloc`" $moviePresetName `"$hbopts`""
    writeLog "HB Command: $hbloc -ArgumentList $cmdArgs`n"
  }

  #Need to keep track of each job indivudually.  Can't run post logic until after job completes.
  #Should also track log files separately as well
  #Run parallel processes
  if ($ParallelProcMax -gt 1 -and $fileCount -gt 1) {
    $logName = $file.baseName + "_" + $timestamp + "_" + ([string]$i).PadLeft($padSize,'0') + "_pp_HBdetails.txt"
    $dtlLogFile = Join-Path -Path $logFilePath -ChildPath $LogName
    Remove-Item $dtlLogFile -ErrorAction SilentlyContinue
    $newProc = Start-Process $hbloc -ArgumentList $cmdArgs -RedirectStandardError $dtlLogFile -PassThru -WindowStyle Minimized
    $procObject = New-Object psobject -Property @{
      id = $newProc.Id
      begTime = $newProc.startTime
      endTime = $null
      begSize = [math]::Round(($file | Measure-Object -Sum Length).Sum / 1GB,3)
      endSize = $null
      dtlLogFile = $dtlLogFile
      baseName = $file.baseName
      fullName = $file.fullName
      newFileName = $newFileName
      countMsg = $countMsg
      processed = $false
    }
    $procList += $procObject
    Start-Sleep -Seconds 1 # Sleep at least one second between submissions
    # Check to see how many are running in the system right now
    $ppcount = @(Get-Process "HandBrakeCLI*" -ErrorAction SilentlyContinue | Select-Object MainWindowTitle | 
            Where-Object MainWindowTitle -like "C:\Program Files\HandBrake\HandBrakeCLI.exe*").Count
    while ($ppcount -ge $ParallelProcMax) {
      writeLog ("Parallel process max limit of " + $ParallelProcMax + " reached. Sleeping for " + $sleepAmt + " seconds`n") -logType "S"
      Start-Sleep -Seconds $sleepAmt
      chkForCompletion($procList) # Perform final processing on any completed jobs
      # Check to see how many are running in the system right now
      $ppcount = @(Get-Process "HandBrakeCLI*" -ErrorAction SilentlyContinue | Select-Object MainWindowTitle | 
              Where-Object MainWindowTitle -like "C:\Program Files\HandBrake\HandBrakeCLI.exe*").Count
    } #while
  } #if ($ParallelProcMax -gt 1 -and $fileCount -gt 1)
  else #Single threaded processing
  {
    $logName = $file.baseName + "_" + $timestamp + "_" + ([string]$i).PadLeft($padSize,'0') + "_HBdetails.txt"
    $dtlLogFile = Join-Path -Path $logFilePath -ChildPath $LogName
    Remove-Item $dtlLogFile -ErrorAction SilentlyContinue
    $newProc = Start-Process $hbloc -ArgumentList $cmdArgs -RedirectStandardError $dtlLogFile -PassThru -Wait -NoNewWindow
    $procObject = New-Object psobject -Property @{
      id = $newProc.Id
      begTime = $newProc.startTime
      endTime = $null
      begSize = [math]::Round(($file | Measure-Object -Sum Length).Sum / 1GB,3)
      endSize = $null
      dtlLogFile = $dtlLogFile
      baseName = $file.baseName
      fullName = $file.fullName
      newFileName = $newFileName
      countMsg = $countMsg
      processed = $false
    }
    $procList += $procObject
    chkForCompletion($procList) # Perform final processing on any completed jobs
  } #else
} #foreach ($file in $videoFiles)

#All of the jobs have now finished being submitted.
#If parallel processing, loop through until all of the jobs have completed
if ($ParallelProcMax -gt 1 -and $fileCount -gt 1) {
  writeLog ("All jobs have been submitted. Now waiting for final jobs to complete`n") -logType "S"
  while (($procList.processed | where-object {$_ -eq $false} | Measure-Object).Count -gt 0) {
    writeLog ("Waiting for jobs to complete - Sleeping " + $sleepAmt + " seconds") -logType "S"
    Start-Sleep -Seconds $sleepAmt 
    chkForCompletion($procList) # Perform final processing on any completed jobs
  } #while
} #if ($ParallelProcMax -gt 1 -and $fileCount -gt 1)
chkForCompletion($procList) # Perform final processing on any completed jobs

# Now that all of the processing is done, process each of the completed sets
writeLog ("HandBrake conversions completed. Now completing final steps")

cleanoldLogs(30) # Remove log files that are more than 30 days old
#writeLog "`nBelow is the list of file(s) containing HandBrake processing details"
foreach ($proc in $procList) {
  $totBegSize += $proc.begSize
  $totEndSize += $proc.endSize
  #$proc.dtlLogFile=$null
  # not sure if there is additional system memory cleanup needed here for completed processes
}

$totSizDiff = ($totBegSize-$totEndSize)
writeLog ("`nDisk space before conversion - " + '{0,7:n3}' -f $totBegSize + " GB")
writeLog (  "Disk space after conversion  - " + '{0,7:n3}' -f $totEndSize + " GB")
#writeLog (  "Amount of disk space saved   - " + '{0,7:n3}' -f $totSizDiff + " GB") -logType "L"
$totSizDiffStr = '{0,7:n3}' -f $totSizDiff + " GB  "
$diskSavingsPCT = [string]([math]::Round(100-($totEndSize / $totBegSize)*100,2)) + "%"
writeLog (  "Amount of disk space saved   - " + $totSizDiffStr + $diskSavingsPCT) -logType "L"
if ($totSizDiff -ge 0) {
  Write-Color "Amount of disk space saved   - ", $totSizDiffStr, $diskSavingsPCT -Color Yellow, Black, Black -BackGroundColor $bgColor, Green, Green }
else {        
  Write-Color "Amount of disk space lost    - ", $totSizDiffStr -Color Yellow, White -BackGroundColor $bgColor, Red }
$endTime  = Get-Date
$timeDiff = getTimeDiff $beginTime $endTime
writeLog ("`nTotal Run time : " + $timeDiff.Hours + " hrs " + $timeDiff.Minutes + 
          " mins " + $timeDiff.Seconds + " secs to process " + $fileCount + " file(s)")
If ($Null -eq (Get-Content $postExecCmd -ErrorAction SilentlyContinue)) {
  if ($postExecCmd) {
    writeLog "`nExecuting requested post execution script"
    writeLog "$postExecCmd $postExecArgs"
    Start-Process $postExecCmd -Wait -ArgumentList $sumLogFile
    writeLog "Post execution completed"
  }
}
WriteLog "`nVidmonHB Processing Completed."
writeLog "`nSummary Log=$sumLogFile`n"
if($errorCount -gt 0) {writeLog "$errorCount Error(s) Found - Please review logs" -logSeverity "E" }
writeLog ""  
Remove-Item $resumeFile -ErrorAction SilentlyContinue
switch ($postLog) {
  "Never" {$null}
  "Error" {if ($errorCount -gt 0) {Invoke-Item $sumLogFile}}
  "Always" {Invoke-Item $sumLogFile}
}
$HistoryObject = New-Object PSObject
$HistoryObject | add-member -membertype NoteProperty -name "yyyy" -value (get-date -Format yyyy)
$HistoryObject | add-member -membertype NoteProperty -name "mm" -value (get-date -Format MM)
$HistoryObject | add-member -membertype NoteProperty -name "dd" -value (get-date -Format dd)
$HistoryObject | add-member -membertype NoteProperty -name "BegSize" -value $totBegSize
$HistoryObject | add-member -membertype NoteProperty -name "EndSize" -value $totEndSize
$HistoryObject | add-member -membertype NoteProperty -name "FileCount" -value $fileCount
$HistoryObject | add-member -membertype NoteProperty -name "ProcessHours" -value $timeDiff.Hours
$HistoryObject | add-member -membertype NoteProperty -name "ProcessMinutes" -value $timeDiff.Minutes
$HistoryObject | add-member -membertype NoteProperty -name "ProcessSeconds" -value $timeDiff.Seconds
$HistoryLogFile = Join-Path -Path $logFilePath -ChildPath "VidMonHB_History.csv"
$HistoryObject | Export-Csv $HistoryLogFile -NoTypeInformation -Append

#--------------------------------------------[Notifications]---------------------------------------------
#If set, send out notification information
$goNotify = $false
if ($postNotify.ToUpper() -eq "ALL") {$goNotify=$true}
if ($postNotify.ToUpper() -eq "ERROR" -and $errorCount -gt 0) {$goNotify=$true}

if ($goNotify -eq $true) {
  <# SMS Message Information
      Alltel   - #1234567890@message.alltel.com
      AT&T     - #1234567890@txt.att.net
      Boost    - #1234567890@myboostmobile.com
      MetroPCS - #1234567890@mymetropcs.com
      Nextel   - #1234567890@messaging.nextel.com
      Sprint   - #1234567890@messaging.sprintpcs.com
      T-Mobile - #1234567890@tmomail.net
      Verizon  - #1234567890@vtext.com
      Virgin   - #1234567890@vmobl.com
  #>

  $textEncoding = [System.Text.Encoding]::UTF8
  # $smsToList array can be a comma delimited list (i.e. @("7327352069@vtext.com", "7325555555@vtext.com")
  $smsToList = @("7327352069@vtext.com")

  if (($errorCount -gt 0) -or $readOnlyErrCnt) {$subject="$serverName - VidMonHB ERROR notification"}
  else {$subject = "$serverName - VidMonHB Successful notification"}

  #Send an SMS alert if the any errors were found
  if($errorCount -gt 0) {
    #$s = New-Object System.Security.SecureString
    #$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "NT AUTHORITY\ANONYMOUS LOGON", $S
    #$smsToList = @()
    foreach ($smsTo in $smsToList) {
      $smsProps = @{
          to         = $smsTo
          From       = $smtpFromEmail
          body       = "Check server"
          subject    = $subject
          #Credential = $creds
          smtpserver = $smtpServer
          }
          Send-MailMessage @smsProps
          Start-Sleep -Seconds 1 # Pause for a few seconds before sending out the next message
    } # Function sendNotification
  }

  #Send an email notification
  #$smtpToEmail="mrpaulwass@hotmail.com"
  $body = "<font face=""verdana"" size=""3"">See attached log for results</font>" 
  if($errorCount -gt 0) {
    $body = "<font face=""verdana"" size=""3""><font color=""red"">HandBrake errors " +
    "occurred during conversion</br></br>"
  }
  if($readOnlyErrCnt) {
    $body += "<font face=""verdana"" size=""3""><font color=""red"">File(s) could not " +
    "be deleted because ReadOnly flag can't be cleared." + 
    " See list below:</font></br></br>"
    foreach ($error in $readOnlyErrCnt) {
      $body += $error
    }
  }  
  $body += "<font color=""black"">See attached log for details</br></br></font>" 

  $emailProps = @{
      From        = $smtpFromEmail
      To          = $smtpToEmail
      Subject     = $subject
      Body        = $body
      SmtpServer  = $smtpServer
      BodyAsHtml  = $true
      Priority    = "High"
      Encoding    = $textEncoding
      #Credential  = $creds
      Attachments = $sumLogFile
      ErrorAction = "Ignore"
      }
  Send-MailMessage @emailProps
} #if ($goNotify -eq $true)

  # if repeatMonitor is set to true, just sit and loop until files are found
  while ($true -eq $repeatMonitor) {
    $date = (get-date).ToString("MM/dd/yyyy hh:mm tt")
    Clear-Host
    Write-Color "`n$date - VidMonHB Monitor mode - Searching for files.  Hit Ctrl-C to break" -BackGroundColor DarkCyan -Color Yellow
    $videoFiles = Get-ChildItem -path $in -Recurse -include $incFiles | where Length -gt $minSize
    $fileCount = ($videofiles | Measure-Object).Count
    if ($fileCount -gt 0) {Invoke-Expression -Command ($PSCommandPath + ' -repeatMonitor $true') ; return }
    Clear-Host
    displayHistory
    $i=60; 
    do {$i--
        $date = (get-date).ToString("MM/dd/yyyy hh:mm tt")
        Write-Progress -Activity "$date - VidMonHB Monitor mode - Will look for new files to process in 1 minute.  Hit Ctrl-C to break" -PercentComplete ($i/60*100) -Status "Seconds remaining $i"
        Start-Sleep -Seconds 1
    } while ($i -ne 0)
    Write-Progress -Activity "$date - VidMonHB Monitor mode - Will look for new files to process in 1 minute.  Hit Ctrl-C to break" -Completed
    Clear-Host
  }

  if ($repeatCtr -eq 0) {return}  #exit script
  else {
    Invoke-Expression -Command ($PSCommandPath + ' -repeatCtr $repeatCtr')
    return }  #exit script
    
<#
  TODO End task any open conhost.exe related to conversion (PID) (may not be necessary)
  TODO Find a better way to calculate ETA (based on prior results??)
  TODO Add health check ffmpeg.exe or handbrakecli
  TODO Assign process to CPU (affinity)
  TODO - Corrections needed for $TVShowBasePath. Needs to have \ as final character.  Chk MovieBasePath too.
  
  $Process = Get-Process Handbrakecli.exe; $Process.ProcessorAffinity=1"
  Core # = Value = BitMask
  Core 1 =   1 = 00000001
  Core 2 =   2 = 00000010
  Core 3 =   4 = 00000100
  Core 4 =   8 = 00001000
  Core 5 =  16 = 00010000
  Core 6 =  32 = 00100000
  Core 7 =  64 = 01000000
  Core 8 = 128 = 10000000
All Cores = 255 = 11111111
Add the decimal values together for which core you want to use. 255 = All 8 cores.
#>
