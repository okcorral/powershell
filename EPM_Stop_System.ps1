
$FileNameDate = (Get-Date).tostring(“yyyyMMdd-hhmmss”)
$FilePrefix = "StopServices"
# $LogPath = “C:\tom\” 
$LogPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent 
$LogPath = $LogPath + “\logs\” 
$LogName = $FilePrefix + “_" + $FileNameDate + “.log”
$FullLogName = $LogPath + $LogName

cls
"FullLogName:  " + $FullLogName

if ($args.Length -gt 0) {
    $ServerList = $args[0].split(",")
} else  {
    # $ServerList = @("DEVEPM01", "DEVEPM02", "DEVEPM03")
    # $ServerList = @("MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")
    #$ServerList = @("MRMEPMRPT4", "MRMEPMRPT3", "MRMEPMRPT2", "MRMEPMRPT1","MRMEPMFND1", "MRMEPMFND0")
    $ServerList = @("ZIRCON")
}

$EPMServiceList = @(
 "beasvc*",
 "OracleProcessManager_ohsInstance*", 
 "HyS9FoundationServices*", 
 "HyS9RaFramework_*",                                                    
 "HyS9RaFrameworkAgent_*"
)


# Add header to logfile
"***************************************************************************************************"   | Out-File  $FullLogName -Append
"Starting EPM Services Query at [$([DateTime]::Now)]."                                                  | Out-File  $FullLogName -Append
"***************************************************************************************************`n" | Out-File  $FullLogName -Append

"Stopping EPM Services...`n" | Out-File $FullLogName -Append

 for ($i=0; $i -lt $ServerList.length; $i++) {

     if ($ServerList[$i] -eq "ZIRCON"){

        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..."
        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
		#Get-Service -Name HyS9RaFrameworkAgent_epmsystem1c -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"						
        #Get-Service -Name HyS9RaFrameworkAgent_epmsystem1b -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
        #Get-Service -Name HyS9RaFrameworkAgent_epmsystem1a -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem1  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
        
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem1 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped" 

        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem1  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance3193331783 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"	

    }

    if ($ServerList[$i] -eq "DEVEPM01"){

        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..."
        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
		#Get-Service -Name HyS9RaFrameworkAgent_epmsystem1c -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"						
        #Get-Service -Name HyS9RaFrameworkAgent_epmsystem1b -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
        #Get-Service -Name HyS9RaFrameworkAgent_epmsystem1a -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem1  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
        
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem1 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped" 

        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem1  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance1649849633 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"	

    }

    if ($ServerList[$i] -eq "DEVEPM02"){

        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..."
        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem2c -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"						
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem2b -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem2a -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem2  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
        
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem2 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped" 

        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem2 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance4217365659 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"	

    }

    if ($ServerList[$i] -eq "MRMEPMFND0"){

        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..."
        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem0  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
        
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem0 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped" 

        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem0 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance357664183 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"	

    }

    if ($ServerList[$i] -eq "MRMEPMFND1"){
        
        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..."
        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem1  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem1 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped" 

        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem1 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance1649849633 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"	

    }

    if ($ServerList[$i] -eq "MRMEPMFND2"){
        
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem2 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped" 

        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Stopping Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem2 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Stopping OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance4217365659 -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"	

    }

    if ($ServerList[$i] -eq "MRMEPMRPT1"){

        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..."
        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem1b  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem1a  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

    }

    if ($ServerList[$i] -eq "MRMEPMRPT2"){

        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..."
        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem2b  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem2a  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

    }

    if ($ServerList[$i] -eq "MRMEPMRPT3"){

        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..."
        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem3b  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem3a  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

    }

    if ($ServerList[$i] -eq "MRMEPMRPT4"){

        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..."
        "*** Stopping IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
		Get-Service -Name HyS9RaFrameworkAgent_epmsystem4b  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem4a  -ComputerName $ServerList[$i] | Set-Service -Status "Stopped"

    }

}


"`nPausing 60 seconds for services shutdown...`n`n"
"`nPausing 60 seconds for services shutdown...`n`n" | Out-File $FullLogName -Append
Start-Sleep -Seconds 60


"Verifying Service Status...`n`n"
"Verifying Service Status...`n`n" | Out-File $FullLogName -Append


for ($i=0; $i -lt $ServerList.length; $i++) {

    "Server: " + $ServerList[$i] | Out-File  $FullLogName -Append 
    Get-Service $EPMServiceList -ComputerName $ServerList[$i] | Sort-Object -Property displayname | Format-Table name, displayname, status -AutoSize | Out-File $FullLogName -Append

}


# Add footer to logfile
"`n***************************************************************************************************" | Out-File  $FullLogName -Append
"EPM Services Query Finished at [$([DateTime]::Now)]."                                                  | Out-File  $FullLogName -Append
"***************************************************************************************************"   | Out-File  $FullLogName -Append

Invoke-Item $FullLogName