# This creates a folder with the date as part of the name
# $folderName = (Get-Date).tostring("dd-MM-yyyy-hh-mm-ss")            
# $folderName = (Get-Date).tostring("yyyyMMdd_hhmm")   
# New-Item -itemType Directory -Path c:\tom -Name $FolderName   
  
# This creates the filename
$FileNameDate = (Get-Date).tostring(“yyyyMMdd-hhmmss”)
$FilePrefix = "log_template"
# $LogPath = “C:\tom\” 
$LogPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent 
$LogPath = $LogPath + “\logs\” 
$LogName = $FilePrefix + “_" + $FileNameDate + “.log”
$FullLogName = $LogPath + $LogName

cls
"FullLogName:  " + $FullLogName

# Add header to logfile
"***************************************************************************************************"   | Out-File  $FullLogName -Append
"Starting EPM Services Query at [$([DateTime]::Now)]."                                                  | Out-File  $FullLogName -Append
"***************************************************************************************************`n" | Out-File  $FullLogName -Append                                                                                               | Out-File  $FullLogName -Append


"This is a sample output line to the log."| Out-File $FullLogName -Append
"This is a sample output line to the log."| Out-File $FullLogName -Append



# Add footer to logfile
"`n***************************************************************************************************" | Out-File  $FullLogName -Append
"EPM Services Query Finished at [$([DateTime]::Now)]."                                                  | Out-File  $FullLogName -Append
"***************************************************************************************************"   | Out-File  $FullLogName -Append

Invoke-Item $FullLogName