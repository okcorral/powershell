# This creates the filename
$FileNameDate = (Get-Date).tostring(“yyyyMMdd-hhmmss”)
$FileName = "Test"
$ext = "csv"
$LogPath = “C:\tom\” 
#$LogPath="\\fsvnode30\epmdev\Output\"
#$LogPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent 
#$LogPath = $LogPath + “\logs\” 

$AppendDate = $true

if ($AppendDate -eq $true){
    $LogName = $FileName + “_" + $FileNameDate + “.” + $ext
} else {
    $LogName = $FileName + “.” + $ext
}
$FullLogName = $LogPath + $LogName

cls
"FullLogName:  " + $FullLogName + "`n"
"Finding info on files in:   " + $dir

#$dir = "\\fsvnode30\epmprod\RAA_DataFiles\root\0000011ee\0000011f\0000013da\0"
#$dir = "\\fsvnode30\epmprod\RAA_DataFiles\root"
#$dir="\\fsvnode30\epmprod\Restore AirSea"
$dir = "C:\Tom\Snagit"

#$logfile="c:\tom\FileSizes.csv"



$f = Get-ChildItem $dir -Recurse | Where { ! $_.PSIsContainer } # | Select  DirectoryName, Name, Length, CreationTime, LastAccessTime # | Export-Csv  -NoTypeInformation $FullLogName
#$f = Get-ChildItem $dir -Recurse | Where { ! $_.PSIsContainer}  and foreach {$_Attributes="DirectoryName","Name"}

foreach ($x in $f){
    
    write-host $x.DirectoryName "," $x.Name  ","  $x.CreationTime  ","  $x.LastAccessTime
    
}