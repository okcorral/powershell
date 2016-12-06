# This creates the filename
$FileNameDate = (Get-Date).tostring(“yyyyMMdd-hhmmss”)
$FilePrefix = "StartServices"
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
    # $ServerList = @("DEVEPM01", "DEVEPM02")
    # $ServerList = @("MRMEPMFND0", "MRMEPMFND1", "MRMEPMFND2", "MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")
    # $ServerList = @("MRMEPMFND0", "MRMEPMFND1", "MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")
    # $ServerList = @("MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")
    $ServerList = @( "ZIRCON")
}

$EPMServiceList = @(
 "OracleProcessManager_ohsInstance*", 
 "HyS9FoundationServices*", 
 "HyS9RaFramework_*",                                                    
 "HyS9RaFrameworkAgent_*"
)


# Add header to logfile
"***************************************************************************************************"   | Out-File  $FullLogName -Append
"Starting EPM Services Query at [$([DateTime]::Now)]."                                                  | Out-File  $FullLogName -Append
"***************************************************************************************************`n" | Out-File  $FullLogName -Append

"Starting EPM Services...`n" | Out-File $FullLogName -Append


 for ($i=0; $i -lt $ServerList.length; $i++) {

     if ($ServerList[$i] -eq "ZIRCON"){

        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance3193331783 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 10

        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem1 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5
        
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem1 -ComputerName $ServerList[$i] | Set-Service -Status "Running"       
        Start-Sleep -Seconds 60

        "*** Starting IR service(s) on " + $ServerList[$i] + " ..."
        "*** Starting IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem1  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        #Get-Service -Name HyS9RaFrameworkAgent_epmsystem1a -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        #Get-Service -Name HyS9RaFrameworkAgent_epmsystem1b -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        #Get-Service -Name HyS9RaFrameworkAgent_epmsystem1c -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5

    }

    if ($ServerList[$i] -eq "DEVEPM01"){

        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance1649849633 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 10

        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem1 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5
        
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem1 -ComputerName $ServerList[$i] | Set-Service -Status "Running"       
        Start-Sleep -Seconds 60

        "*** Starting IR service(s) on " + $ServerList[$i] + " ..."
        "*** Starting IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem1  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        #Get-Service -Name HyS9RaFrameworkAgent_epmsystem1a -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        #Get-Service -Name HyS9RaFrameworkAgent_epmsystem1b -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        #Get-Service -Name HyS9RaFrameworkAgent_epmsystem1c -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5

    }
 
    if ($ServerList[$i] -eq "DEVEPM02"){

        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance4217365659 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 10

        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem2 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5
        
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem2 -ComputerName $ServerList[$i] | Set-Service -Status "Running"       
        Start-Sleep -Seconds 60

        "*** Starting IR service(s) on " + $ServerList[$i] + " ..."
        "*** Starting IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem2  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem2a -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem2b -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem2c -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5

    }

    if ($ServerList[$i] -eq "MRMEPMFND0"){

        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance357664183 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 10

        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem0 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5
        
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem0 -ComputerName $ServerList[$i] | Set-Service -Status "Running"       
        Start-Sleep -Seconds 60

        "*** Starting IR service(s) on " + $ServerList[$i] + " ..."
        "*** Starting IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem0  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5

    }

    if ($ServerList[$i] -eq "MRMEPMFND1"){

        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance1649849633 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 10

        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem1 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5
        
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem1 -ComputerName $ServerList[$i] | Set-Service -Status "Running"       
        Start-Sleep -Seconds 60

        "*** Starting IR service(s) on " + $ServerList[$i] + " ..."
        "*** Starting IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem1  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5

    }

    if ($ServerList[$i] -eq "MRMEPMFND2"){

        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..."
        "*** Starting OHTTP (OHS) Web Server on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name OracleProcessManager_ohsInstance4217365659 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 10

        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Foundation Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9FoundationServices_epmsystem2 -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5
        
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..."
        "*** Starting Managed Server Framework Services on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFramework_epmsystem2 -ComputerName $ServerList[$i] | Set-Service -Status "Running"       
        Start-Sleep -Seconds 60

    }

    if ($ServerList[$i] -eq "MRMEPMRPT1"){

        "*** Starting IR service(s) on " + $ServerList[$i] + " ..."
        "*** Starting IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem1a  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem1b  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5

    }

    if ($ServerList[$i] -eq "MRMEPMRPT2"){

        "*** Starting IR service(s) on " + $ServerList[$i] + " ..."
        "*** Starting IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem2a  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem2b  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5

    }

    if ($ServerList[$i] -eq "MRMEPMRPT3"){

        "*** Starting IR service(s) on " + $ServerList[$i] + " ..."
        "*** Starting IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem3a  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem3b  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5

    }

    if ($ServerList[$i] -eq "MRMEPMRPT4"){

        "*** Starting IR service(s) on " + $ServerList[$i] + " ..."
        "*** Starting IR service(s) on " + $ServerList[$i] + " ..." | Out-File  $FullLogName -Append
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem4a  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Get-Service -Name HyS9RaFrameworkAgent_epmsystem4b  -ComputerName $ServerList[$i] | Set-Service -Status "Running"
        Start-Sleep -Seconds 5

    }

}


"`nPausing 60 seconds for services startup...`n`n"
"`nPausing 60 seconds for services startup...`n`n" | Out-File $FullLogName -Append
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