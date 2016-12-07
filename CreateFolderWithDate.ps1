#$folderName = (Get-Date).tostring("dd-MM-yyyy-hh-mm-ss")            
$folderName = (Get-Date).tostring("yyyyMMdd_hhmm")        
New-Item -itemType Directory -Path c:\tom -Name $FolderName

$FileName = (Get-Date).tostring(“yyyyMMdd-hhmmss”)
New-Item -itemType File -Path c:\tom -Name ($FileName + “.log”)