 
#### Spreadsheet Location
 $DirectoryToSaveTo = "c:\tom"
 $date=Get-Date -format "yyyy-MM-d"
 $Filename="serverinfo-$($date)"
 $FromEmail="<ToEmail>"
 $ToEmail="<FromEmail>"
 $SMTPMail="<SMTP MAIL>"
 
###InputLocation
# $Computers = Get-Content "c:\tom\server.txt"
# check single server
#$Computers = @("MRMEPMFND0")
# check Hyperion 9.3 servers
# $Computers = @("DEVHYPR01", "DEVHYPR02", "MRMHYPR02", "MRMHYPR03", "MRMHYPR04", "MRMHYPR05", "MRMHYPR06")
# check OBIEE servers
# $Computers = @("DEVOBI02", "MRMOBI02")
# check EPM 11 servers
 $Computers = @( "MRMEPMFND0", "MRMEPMFND1",  "MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")
# check All BI servers

#$Computers = @("MRMEPM08",
#               "DEVEPM01", "DEVEPM02",  "DEVEPM03", 
#               "MRMEPMFND0", "MRMEPMFND1", "MRMEPMFND2", "MRMEPMRPT1", "MRMEPMRPT2", "MRMEPMRPT3", "MRMEPMRPT4")

 
#$Computers = @("MRMSSVCR01")

# before we do anything else, are we likely to be able to save the file?
# if the directory doesn't exist, then create it
if (!(Test-Path -path "$DirectoryToSaveTo")) #create it if not existing
  {
  New-Item "$DirectoryToSaveTo" -type directory | out-null
  }
  


#Create a new Excel object using COM 
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $True
$Excel = $Excel.Workbooks.Add()
$Excel.Worksheets.Add()
$Sheet = $Excel.Worksheets.Item(1)

$sheet.Name = 'Server Inventory'
#Create a Title for the first worksheet
$row = 1
$Column = 1
$Sheet.Cells.Item($row,$column)= 'Server Inventory'

$range = $Sheet.Range("a1","s2")
$range.Merge() | Out-Null
$range.VerticalAlignment = -4160

#Give it a nice Style so it stands out
$range.Style = 'Title'

#Increment row for next set of data
$row++;$row++

#Save the initial row so it can be used later to create a border
#Counter variable for rows
$intRow = $row
$xlOpenXMLWorkbook=[int]51

#Read thru the contents of the server.txt file

$Sheet.Cells.Item($intRow,1)  ="Name"
$Sheet.Cells.Item($intRow,2)  ="status"
$Sheet.Cells.Item($intRow,3)  ="OS"
$Sheet.Cells.Item($intRow,4)  ="Domain Role"
$Sheet.Cells.Item($intRow,5)  ="ProcessorName"
$Sheet.Cells.Item($intRow,6)  ="Manufacturer"
$Sheet.Cells.Item($intRow,7)  ="Model"
$Sheet.Cells.Item($intRow,8)  ="SystemType"
$Sheet.Cells.Item($intRow,9)  ="Last Boot Time"
$Sheet.Cells.Item($intRow,10) ="Bios Version"
$Sheet.Cells.Item($intRow,11) ="CPU Info"
$Sheet.Cells.Item($intRow,12) ="NoOfProcessors"
$Sheet.Cells.Item($intRow,13) ="Total Physical Memory"
$Sheet.Cells.Item($intRow,14) ="Total Free Physical Memory"
$Sheet.Cells.Item($intRow,15) ="Total Virtual Memory"
$Sheet.Cells.Item($intRow,16) ="Total Free Virtual Memory"
$Sheet.Cells.Item($intRow,17) ="Disk Info"
$Sheet.Cells.Item($intRow,18) ="FQDN"
$Sheet.Cells.Item($intRow,19) ="IPAddress"

for ($col = 1; $col –le 19; $col++)
     {
          $Sheet.Cells.Item($intRow,$col).Font.Bold = $True
          $Sheet.Cells.Item($intRow,$col).Interior.ColorIndex = 48
          $Sheet.Cells.Item($intRow,$col).Font.ColorIndex = 34
     }

$intRow++


Function GetStatusCode
{ 
	Param([int] $StatusCode)  
	switch($StatusCode)
	{
		0 		{"Success"}
		11001   {"Buffer Too Small"}
		11002   {"Destination Net Unreachable"}
		11003   {"Destination Host Unreachable"}
		11004   {"Destination Protocol Unreachable"}
		11005   {"Destination Port Unreachable"}
		11006   {"No Resources"}
		11007   {"Bad Option"}
		11008   {"Hardware Error"}
		11009   {"Packet Too Big"}
		11010   {"Request Timed Out"}
		11011   {"Bad Request"}
		11012   {"Bad Route"}
		11013   {"TimeToLive Expired Transit"}
		11014   {"TimeToLive Expired Reassembly"}
		11015   {"Parameter Problem"}
		11016   {"Source Quench"}
		11017   {"Option Too Big"}
		11018   {"Bad Destination"}
		11032   {"Negotiating IPSEC"}
		11050   {"General Failure"}
		default {"Failed"}
	}
}


Function GetUpTime
{
	param([string] $LastBootTime)
	$Uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
	"Days: $($Uptime.Days); Hours: $($Uptime.Hours); Minutes: $($Uptime.Minutes); Seconds: $($Uptime.Seconds)" 
}

    
	



foreach ($Computer in $Computers)
 {

 TRY {
 $OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer
 $Bios = Get-WmiObject -Class Win32_BIOS -ComputerName $Computer
 $sheetS = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer
 $sheetPU = Get-WmiObject -Class Win32_Processor -ComputerName $Computer
 $drives = Get-WmiObject -ComputerName $Computer Win32_LogicalDisk | Where-Object {$_.DriveType -eq 3}
 $pingStatus = Get-WmiObject -Query "Select * from win32_PingStatus where Address='$Computer'"
 $IPAddress=(Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer | ? {$_.IPEnabled}).ipaddress
 $FQDN=[System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
 $OSRunning = $OS.caption + " " + $OS.OSArchitecture + " SP " + $OS.ServicePackMajorVersion
 $NoOfProcessors=$sheetS.numberofProcessors
 $name=$SheetPU|select name -First 1
 $Manufacturer=$sheetS.Manufacturer
 $Model=$sheetS.Model
 $systemType=$sheetS.SystemType
 $ProcessorName=$SheetPU|select name -First 1
 $DomainRole = $sheetS.DomainRole
 $TotalAvailMemory = $OS.totalvisiblememorysize/1kb
 $TotalVirtualMemory = $OS.totalvirtualmemorysize/1kb
 $TotalFreeMemory = $OS.FreePhysicalMemory/1kb
 $TotalFreeVirtualMemory = $OS.FreeVirtualMemory/1kb
 $TotalMem = "{0:N2}" -f $TotalAvailMemory
 $TotalVirt = "{0:N2}" -f $TotalVirtualMemory
 $FreeMem = "{0:N2}" -f $TotalFreeMemory
 $FreeVirtMem = "{0:N2}" -f $TotalFreeVirtualMemory
 $date = Get-Date
 $uptime = $OS.ConvertToDateTime($OS.lastbootuptime)
 $BiosVersion = $Bios.Manufacturer + " " + $Bios.SMBIOSBIOSVERSION + " " + $Bios.ConvertToDateTime($Bios.Releasedate)
 $sheetPUInfo = $name.Name + " & has " + $sheetPU.NumberOfCores + " Cores & the FSB is " + $sheetPU.ExtClock + " Mhz"
 $sheetPULOAD = $sheetPU.LoadPercentage
 
 if($pingStatus.StatusCode -eq 0)
	{
		$Status = GetStatusCode( $pingStatus.StatusCode )
    }
else
	{
	$Status = GetStatusCode( $pingStatus.StatusCode )
   	}
	
	
 if (($DomainRole -eq "0") -or ($DomainRole -eq "1"))
 {
 $Role = "Work Station"
 }
 elseif (($DomainRole -eq "2") -or ($DomainRole -eq "3"))
 {
 $Role = "Member Server"
 }
 elseif (($DomainRole -eq "4") -or ($DomainRole -eq "5"))
 {
 $Role = "Domain Controller"
 }
 else
 {
 $Role = "Unknown"
 }
 }
 CATCH
 {
 $pcnotfound = "true"
 }
 #### Pump Data to Excel
 if ($pcnotfound -eq "true")
 {
 $sheet.Cells.Item($intRow, 1) = "$($computer) Not Found "
 }
 else
 {
 $sheet.Cells.Item($intRow, 1) = $computer
 $sheet.Cells.Item($intRow, 2) = $status
 $sheet.Cells.Item($intRow, 3) = $OSRunning
 $sheet.Cells.Item($intRow, 4) = $Role
 $sheet.Cells.Item($intRow, 5) = $name.name
 $Sheet.Cells.Item($intRow, 6) = $Manufacturer
 $Sheet.Cells.Item($intRow, 7) = $Model
 $Sheet.Cells.Item($intRow, 8) = $SystemType
 $sheet.Cells.Item($intRow, 9) = $uptime
 $sheet.Cells.Item($intRow, 10)= $BiosVersion
 $sheet.Cells.Item($intRow, 11)= $sheetPUInfo
 $sheet.Cells.Item($intRow, 12)=$NoOfProcessors
 $sheet.Cells.Item($intRow, 13)= "$TotalMem MB"
 $sheet.Cells.Item($intRow, 14)= "$FreeMem MB"
 $sheet.Cells.Item($intRow, 15)= "$TotalVirt MB"
 $sheet.Cells.Item($intRow, 16)= "$FreeVirtMem MB"
 $sheet.Cells.Item($intRow, 19)=$IPAddress
 $sheet.Cells.Item($intRow, 18)=$FQDN

 
$driveStr = ""
 foreach($drive in $drives)
 {
 $size1 = $drive.size / 1GB
 $size = "{0:N2}" -f $size1
 $free1 = $drive.freespace / 1GB
 $free = "{0:N2}" -f $free1
 $freea = $free1 / $size1 * 100
 $freeb = "{0:N2}" -f $freea
 $ID = $drive.DeviceID
 $driveStr += "$ID = Total Space: $size GB / Free Space: $free GB / Free (Percent): $freeb % ` "
 }
 $sheet.Cells.Item($intRow, 17) = $driveStr
 }

 
$intRow = $intRow + 1
 $pcnotfound = "false"
 }

$erroractionpreference = “SilentlyContinue” 

$Sheet.UsedRange.EntireColumn.AutoFit()
########################################333

$Sheet = $Excel.Worksheets.Item(2)
$sheet.Name = 'DiskSpace'
$Sheet.Activate() | Out-Null

#Create a Title for the first worksheet
$row = 1
$Column = 1
$Sheet.Cells.Item($row,$column)= 'Disk Space Information'

$range = $Sheet.Range("a1","h2")
$range.Merge() | Out-Null
$range.VerticalAlignment = -4160

#Give it a nice Style so it stands out
$range.Style = 'Title'

#Increment row for next set of data
$row++;$row++

#Save the initial row so it can be used later to create a border
$initalRow = $row

#Create a header for Disk Space Report; set each cell to Bold and add a background color
$Sheet.Cells.Item($row,$column)= 'Computername'
$Sheet.Cells.Item($row,$column).Interior.ColorIndex =48
$Sheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$Sheet.Cells.Item($row,$column)= 'DeviceID'
$Sheet.Cells.Item($row,$column).Interior.ColorIndex =48
$Sheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$Sheet.Cells.Item($row,$column)= 'VolumeName'
$Sheet.Cells.Item($row,$column).Interior.ColorIndex =48
$Sheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$Sheet.Cells.Item($row,$column)= 'TotalSizeGB'
$Sheet.Cells.Item($row,$column).Interior.ColorIndex =48
$Sheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$Sheet.Cells.Item($row,$column)= 'UsedSpaceGB'
$Sheet.Cells.Item($row,$column).Interior.ColorIndex =48
$Sheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$Sheet.Cells.Item($row,$column)= 'FreeSpaceGB'
$Sheet.Cells.Item($row,$column).Interior.ColorIndex =48
$Sheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$Sheet.Cells.Item($row,$column)= '%Free'
$Sheet.Cells.Item($row,$column).Interior.ColorIndex =48
$Sheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$Sheet.Cells.Item($row,$column)= 'State'
$Sheet.Cells.Item($row,$column).Interior.ColorIndex =48
$Sheet.Cells.Item($row,$column).Font.Bold=$True

#Set up a header filter
$headerRange = $Sheet.Range("a3","h3")
$headerRange.AutoFilter() | Out-Null

#Increment Row and reset Column back to first column
$row++
$Column = 1
$critical=0
$warning=0
$good=0

#Get the drives and filter out CD/DVD drives
foreach ($computer in $Computers)
 {
$diskDrives = Get-WmiObject win32_LogicalDisk -Filter "DriveType='3'" -ComputerName $computer

#Process each disk in the collection and write to spreadsheet
ForEach ($disk in $diskDrives) {
    $Sheet.Cells.Item($row,1)= $disk.__Server
    $Sheet.Cells.Item($row,2)= $disk.DeviceID
    $Sheet.Cells.Item($row,3)= $disk.VolumeName
    $Sheet.Cells.Item($row,4)= [math]::Round(($disk.Size /1GB),2)
    $Sheet.Cells.Item($row,5)= [math]::Round((($disk.Size - $disk.FreeSpace)/1GB),2)
    $Sheet.Cells.Item($row,6)= [math]::Round(($disk.FreeSpace / 1GB),2)
    $Sheet.Cells.Item($row,7)= ("{0:P}" -f ($disk.FreeSpace / $disk.Size))  
   
    #Determine if disk needs to be flagged for warning or critical alert
    If ($disk.FreeSpace -lt 5GB -AND ("{0:P}" -f ($disk.FreeSpace / $disk.Size))  -lt 40) {
        $Sheet.Cells.Item($row,8) = "Critical"
        $critical++
        #Check to see if space is near empty and use appropriate background colors
        $range = $Sheet.Range(("A{0}"  -f $row),("H{0}"  -f $row))
        $range.Select() | Out-Null    
        #Critical threshold         
        $range.Interior.ColorIndex = 3
    } ElseIf ($disk.FreeSpace -lt 10GB -AND ("{0:P}" -f ($disk.FreeSpace / $disk.Size)) -lt 60) {
        $Sheet.Cells.Item($row,8) = "Warning"
        $range = $Sheet.Range(("A{0}"  -f $row),("H{0}"  -f $row))
        $range.Select() | Out-Null        
        $warning++
        $range.Interior.ColorIndex = 6
    } Else {
        $Sheet.Cells.Item($row,8) = "Good"
        $good++
    }

     $row++
}
}

#Add a border for data cells
$row--
$dataRange = $Sheet.Range(("A{0}"  -f $initalRow),("H{0}"  -f $row))
7..12 | ForEach {
    $dataRange.Borders.Item($_).LineStyle = 1
    $dataRange.Borders.Item($_).Weight = 2
}

#Auto fit everything so it looks better

$usedRange = $Sheet.UsedRange															
$usedRange.EntireColumn.AutoFit() | Out-Null

$critical
$warning
$good

$sheet = $excel.Worksheets.Item(2) 
 
$row++;$row++

<#

$beginChartRow = $Row

$Sheet.Cells.Item($row,$Column) = 'Critical'
$Column++
$Sheet.Cells.Item($row,$Column) = 'Warning'
$Column++
$Sheet.Cells.Item($row,$Column) = 'Good'
$Column = 1
$row++
#Critical formula
$Sheet.Cells.Item($row,$Column)=$critical
$Column++
#Warning formula
$Sheet.Cells.Item($row,$Column)=$warning
$Column++
#Good formula
$Sheet.Cells.Item($row,$Column)= $good

$endChartRow = $row

$chartRange = $Sheet.Range(("A{0}" -f $beginChartRow),("C{0}" -f $endChartRow))

##Add a chart to the workbook
#Open a sheet for charts
$temp = $sheet.Charts.Add()
$temp.Delete()
$chart = $sheet.Shapes.AddChart().Chart
$sheet.Activate()

#Configure the chart
##Use a 3D Pie Chart
$chart.ChartType = 70
$chart.Elevation = 40
#Give it some color
$sheet.Shapes.Item("Chart 1").Fill.ForeColor.TintAndShade = .34
$sheet.Shapes.Item("Chart 1").Fill.ForeColor.ObjectThemeColor = 5
$sheet.Shapes.Item("Chart 1").Fill.BackColor.TintAndShade = .765
$sheet.Shapes.Item("Chart 1").Fill.ForeColor.ObjectThemeColor = 5

$sheet.Shapes.Item("Chart 1").Fill.TwoColorGradient(1,1)

#Set the location of the chart
$sheet.Shapes.Item("Chart 1").Placement = 3
$sheet.Shapes.Item("Chart 1").Top = 30
$sheet.Shapes.Item("Chart 1").Left = 600

$chart.SetSourceData($chartRange)
$chart.HasTitle = $True

$chart.ApplyLayout(6,69)
$chart.ChartTitle.Text = "Disk Space Report"
$chart.ChartStyle = 26
$chart.PlotVisibleOnly = $False
$chart.SeriesCollection(1).DataLabels().ShowValue = $True
$chart.SeriesCollection(1).DataLabels().Separator = ("{0}" -f [char]10)

$chart.SeriesCollection(1).DataLabels().Position = 2
#Critical
$chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = 255
#Warning
$chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = 65535
#Good
$chart.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = 5287936

#Hide the data
$chartRange.EntireRow.Hidden = $True

#>

$sheet.Name = 'DiskInformation'




$filename = "$DirectoryToSaveTo\$filename.xlsx"
if (test-path $filename ) { rm $filename } #delete the file if it already exists
$Sheet.UsedRange.EntireColumn.AutoFit()
$Excel.SaveAs($filename, $xlOpenXMLWorkbook) #save as an XML Workbook (xslx)
$Excel.Saved = $True
$Excel.Close()
$Excel.DisplayAlerts = $False
$Excel.quit()


Function sendEmail([string]$emailFrom, [string]$emailTo, [string]$subject,[string]$body,[string]$smtpServer,[string]$filePath)
{
#initate message
$email = New-Object System.Net.Mail.MailMessage 
$email.From = $emailFrom
$email.To.Add($emailTo)
$email.Subject = $subject
$email.Body = $body
# initiate email attachment 
$emailAttach = New-Object System.Net.Mail.Attachment $filePath
$email.Attachments.Add($emailAttach) 
#initiate sending email 
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($email)
}

#Call Function 

$message = @" 
Hi Team,

The Discovery of Windows Server and Disk Space information for all the listed servers.

Autogenerated Email!!! Please do not reply.

Thank you, 
xyz.com

"@        
$date=get-date

#sendEmail -emailFrom $fromEmail -emailTo $ToEmail -subject "Windows Server Inventory & Disk Details -$($date)" -body $message -smtpServer $SMTPMail -filePath $filename

Invoke-Item $filename