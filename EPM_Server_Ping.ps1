# This creates the filename
$FileNameDate = (Get-Date).tostring(“MM/dd/yyyy hh:mm:ss”)
$hostname = hostname
$FilePrefix = $hostname  + "_EPM_Server_Ping"
# $LogPath = “C:\tom\” 
$LogPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent 
$LogPath = $LogPath + “\netlogs\” 
#$LogName = $FilePrefix + “_" + $FileNameDate + “.log”
$LogName = $FilePrefix + “.log”
$FullLogName = $LogPath + $LogName

cls
"FullLogName:  " + $FullLogName

#$servers = @("MRMEPMFND0", "MRMEPMFND1", "MRMEPMFND2", "MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")  
$servers = @("MRMEPMFND0")  
  
foreach ( $server in $servers ) { 
    if ($server -ne $hostname){ 
          
        if ((test-Connection -ComputerName $server -Count 2 -Quiet) -eq $true ) {           

            "`n" + $server + ":   Checking Network Connectivity at " + $FileNameDate   | Out-File  $FullLogName -Append   
            test-Connection -ComputerName $server -Count 2 -Verbose                    | Out-File  $FullLogName -Append                          
        } else {   
                                                                
            "`"Computer $server not Pinging, i am going to do traceroute now.`" `n`n"  | Out-File  $FullLogName -Append                          
        }  
    } 
}