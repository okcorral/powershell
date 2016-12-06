# This creates the filename
$FileNameDate = (Get-Date).tostring(“yyyyMMdd-hhmmss”)
$FilePrefix = "QueryProcesses"
# $LogPath = “C:\tom\” 
$LogPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent 
$LogPath = $LogPath + “\logs\” 
$LogName = $FilePrefix + “_" + $FileNameDate + “.log”
$FullLogName = $LogPath + $LogName

cls
"FullLogName:  " + $FullLogName

# Add header to logfile
"***************************************************************************************************"   | Out-File  $FullLogName -Append
"Starting EPM Processes Query at [$([DateTime]::Now)]."                                                 | Out-File  $FullLogName -Append
"***************************************************************************************************`n" | Out-File  $FullLogName -Append
      

$FoundationServerList = @("ZIRCON")
$ReportingServerList = @("ZIRCON")

#$FoundationServerList =@("MRMEPMFND0", "MRMEPMFND1")
#$ReportingServerList = @("MRMEPMFND0", "MRMEPMFND1", "MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")


"Foundation Services/Processes:"  | Out-File  $FullLogName -Append 

for ($i=0; $i -lt $FoundationServerList.length; $i++) {

     get-process -name Apache*, HyS9FoundationServices*,  HyS9RaFramework_*, opmn* `
                 -computername  $FoundationServerList[$i]  | Sort-Object -Descending -Property WS |  format-table -property  `
                        MachineName, ProcessName, Id,
                        @{Label="Memory(K)";Expression={[int]($_.WS/1024)}} -AutoSize `
                        | Out-File $FullLogName -Append -Width 10000
           
}


"Reporting Services/Processes:"  | Out-File  $FullLogName -Append 

for ($i=0; $i -lt $ReportingServerList.length; $i++) {

     get-process -name java*, BIService*,das*, HyS9RaFrameworkAgent_* `
                 -computername  $ReportingServerList[$i]  | Sort-Object -Descending -Property WS | format-table -property  `
                        MachineName, ProcessName, Id,
                        @{Label="Memory(K)";Expression={[int]($_.WS/1024)}} -AutoSize | Out-File $FullLogName -Append -Width 10000
           
}


# Add footer to logfile
"`n***************************************************************************************************" | Out-File  $FullLogName -Append
"EPM Processes Query Finished at [$([DateTime]::Now)]."                                                 | Out-File  $FullLogName -Append
"***************************************************************************************************"   | Out-File  $FullLogName -Append

Invoke-Item $FullLogName