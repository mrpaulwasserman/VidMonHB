# HandBrake Batch Converter - Windows PowerShell Script

This script will find specified types of video files and process them
through the HandBrake Command Line Interface.  Parameters allow script
to define functionality such as input folders, output folders, type of
video files to search for, etc.  

  - This script will handle up to 9999 files at a time  
  - Script will attempt to parse name to determine output folder structure  
 
Sample Output Folder Below:  
- D:\Media\TV Show Name\Season 01\TV Show Name - S01E01 - Episode Name  

\Media\00. Movies is the based default location:

\Media
|-- 00. Movies 
|   |-- 01. Movies 1920-1979
|   |-- 02. Movies 1981-1999
|   |-- 03. Movies 2000-2019
|   |-- 04. Movies 2020-2029


\Media\20. TV Shows is the base default location:

\Media\20. TV Shows
|-- 20. TV Shows
|   |-- Dynasty
|   |-- Green Acres
|   |-- The Andy Griffith Show
|   |-- The Flintstones

------------
### PowerShell Modules Needed
First time execution may require running the following command (for PowerShell 5 & lower)
-  ##### Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force

If you want to recycle files (instead of delete), install this Powershell module
-  ##### Install-Module -Name Recycle -RequiredVersion 1.0.2 -Scope CurrentUser -Force

------------
File metadata (specifically Title) will attempt to be removed if the
taglib-sharp.dll module has been downloaded and placed in the same folder
with this script.
- Note : This is not required because HandBrake will typically not transfer          metadata during the conversion process.

## Links to additional programs
HandBrakeCLI - Command Line Interface version of HandBrake video converter  
 - Use the following link to download the HandBrakeCLI.exe program  
  https://handbrake.fr/downloads2.php

------------
## The following MetaData modification programs are not required.

TagLib-Sharp - Used for updating Title metadata
- Use the following link to download the taglib-sharp.dll (not required)  
  https://www.dllme.com/dll/files/taglib-sharp_dll.html

Alternative to TagLib is exifTool (can be added as post execution step)
- Use the following link to download the exitTool.exe (not required)  
  https://exiftool.org/
