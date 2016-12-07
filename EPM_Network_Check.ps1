# This creates the filename
$FileNameDate = (Get-Date).tostring(“yyyy.MM.dd_hhmm”)
$hostname = hostname
$FilePrefix = $hostname  + "_EPM_Network_Check"
# $LogPath = “C:\tom\” 
$LogPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent 
$LogPath = $LogPath + “\netlogs\” 
#$LogName = $FilePrefix + “_" + $FileNameDate + “.log”
$LogName = $FilePrefix + “.log”
$FullLogName = $LogPath + $LogName

cls
"FullLogName:  " + $FullLogName


# Add header to logfile
"***************************************************************************************************"   | Out-File  $FullLogName -Append
"Starting EPM Network Check at [$([DateTime]::Now)]."                                                   | Out-File  $FullLogName -Append
"***************************************************************************************************`n" | Out-File  $FullLogName -Append

"Starting EPM Network Check...`n`n" | Out-File $FullLogName -Append


#$servers = @("MRMEPMFND0", "MRMEPMFND1", "MRMEPMFND2", "MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")  
$servers = @("MRMEPMFND0")  
  
foreach ( $server in $servers ) { 
if ($server -ne $hostname){ 
          
            if ((test-Connection -ComputerName $server -Count 2 -Quiet) -eq $true ) { 
          
            "`n========================= Starting Ping for $server ========== "        | Out-File  $FullLogName -Append 
            "`n" + $server + ":   Checking Network Connectivity`n "                    | Out-File  $FullLogName -Append   
            test-Connection -ComputerName $server -Count 2 -Verbose                    | Out-File  $FullLogName -Append  
            $server + " is alive and Pinging. " | Out-File  $FullLogName -Append
            "========================= Ping for $server Done ============== `n"        | Out-File  $FullLogName -Append 
                    
            "`n========================= Starting Traceroute for $server ========== "  | Out-File  $FullLogName -Append 
            tracert -d  $server                                                        | Out-File  $FullLogName -Append 
            "========================= Traceroute for $server Done ============== `n"  | Out-File  $FullLogName -Append 
              
          
                    } else {   
                                          
                      
                    "`"Computer $server not Pinging, i am going to do traceroute now.`" `n`n"  | Out-File  $FullLogName -Append   
          
                    "========================= Starting Traceroute for $server ========== `n"  | Out-File  $FullLogName -Append  
                    tracert -d  $server                                                        | Out-File  $FullLogName -Append 
                    "========================= Traceroute for $server Done ============== `n"  | Out-File  $FullLogName -Append   
              
                    }  
} 
}
 
"`n========================= Testing Done for all Servers ============== `n"  | Out-File  $FullLogName -Append  
"`n"  
"`n"  
                      
# Add footer to logfile
"`n***************************************************************************************************" | Out-File  $FullLogName -Append
"EPM Network Check Finished at [$([DateTime]::Now)]."                                                  | Out-File  $FullLogName -Append
"***************************************************************************************************"   | Out-File  $FullLogName -Append