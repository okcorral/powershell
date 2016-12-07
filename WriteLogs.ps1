# This creates the filename
$FileNameDate = (Get-Date).tostring(“yyyyMMdd-hhmmss”)
$FileName = "Test"
$ext = "log"
# $LogPath = “C:\tom\” 
$LogPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent 
$LogPath = $LogPath + “\logs\” 

$AppendDate = $true


if ($AppendDate -eq $true){
    $LogName = $FileName + “_" + $FileNameDate + “.” + $ext
} else {
    $LogName = $FileName + “.” + $ext
}
$FullLogName = $LogPath + $LogName

cls
"FullLogName:  " + $FullLogName + "`n"

$ServerList = @("DEVEPM01")


$msg = "this is a test on $ServerList ..."
write-host $msg
$msg | out-file  $FullLogName -Append

$msg = "`n"
write-host $msg
$msg | Out-File  $FullLogName -Append

Invoke-Item $FullLogName