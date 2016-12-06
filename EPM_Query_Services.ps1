# This creates a folder with the date as part of the name
# $folderName = (Get-Date).tostring("dd-MM-yyyy-hh-mm-ss")            
# $folderName = (Get-Date).tostring("yyyyMMdd_hhmm")   
# New-Item -itemType Directory -Path c:\tom -Name $FolderName   
  
# This creates the filename
$FileNameDate = (Get-Date).tostring(“yyyyMMdd-hhmmss”)
$FilePrefix = "QueryServices"
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
"***************************************************************************************************`n" | Out-File  $FullLogName -Append
                                                                                                

if ($args.Length -gt 0) {
    $ServerList = $args[0].split(",")
} else  {
    # $ServerList = @("DEVEPM01", "DEVEPM02")
    # $ServerList = @("MRMEPMFND0", "MRMEPMFND1", "MRMEPMFND2", "MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")
    #$ServerList = @("MRMEPMFND0", "MRMEPMFND1", "MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")
    $ServerList = @("ZIRCON")
}

$EPMServiceList = @(
 "OracleProcessManager_ohsInstance*", 
 "HyS9FoundationServices*", 
 "HyS9RaFramework_*",                                                    
 "HyS9RaFrameworkAgent_*"
)


for ($i=0; $i -lt $ServerList.length; $i++) {

    "Server: " + $ServerList[$i] | Out-File  $FullLogName -Append 
    Get-Service $EPMServiceList -ComputerName $ServerList[$i] | Sort-Object -Property displayname | Format-Table name, displayname, status -AutoSize | Out-File $FullLogName -Append -Width 10000

}


# Add footer to logfile
"`n***************************************************************************************************" | Out-File  $FullLogName -Append
"EPM Services Query Finished at [$([DateTime]::Now)]."                                                  | Out-File  $FullLogName -Append
"***************************************************************************************************"   | Out-File  $FullLogName -Append

Invoke-Item $FullLogName