#$dir = "\\fsvnode30\epmprod\RAA_DataFiles\root\0000011ee\0000011f\0000013da\0"
$dir = "\\fsvnode30\epmprod\RAA_DataFiles\root"
#$dir="\\fsvnode30\epmprod\Restore AirSea"
# $dir = "\\mrmepm08\C$\Windows"
#$dir = "C:\Tom\ocebackup\.Connections"
#$dir = "\\mrmssvcr01\f$"

#$logfile="c:\tom\FileSizes.csv"

$logfile="\\fsvnode30\epmdev\Output\test_fileinfo2.csv"

Get-ChildItem $dir -Recurse | Where { ! $_.PSIsContainer } | Select  DirectoryName, Name, Length, CreationTime, LastAccessTime | Export-Csv  -NoTypeInformation $logfile 