$Instance = "4a"

Write-Host "Killing Process Tree on instance $Instance :"

switch ($Instance) {

    "1a" { taskkill.exe /s MRMEPMRPT1 /im HyS9RaFrameworkAgent_epmsystem1a.exe /t /F }
    "1b" { taskkill.exe /s MRMEPMRPT1 /im HyS9RaFrameworkAgent_epmsystem1b.exe /t /F }

    "2a" { taskkill.exe /s MRMEPMRPT2 /im HyS9RaFrameworkAgent_epmsystem2a.exe /t /F }
    "2b" { taskkill.exe /s MRMEPMRPT2 /im HyS9RaFrameworkAgent_epmsystem2b.exe /t /F }

    "3a" { taskkill.exe /s MRMEPMRPT3 /im HyS9RaFrameworkAgent_epmsystem3a.exe /t /F }
    "3b" { taskkill.exe /s MRMEPMRPT3 /im HyS9RaFrameworkAgent_epmsystem3b.exe /t /F }

    "4a" { taskkill.exe /s MRMEPMRPT4 /im HyS9RaFrameworkAgent_epmsystem4a.exe /t /F }
    "4b" { taskkill.exe /s MRMEPMRPT4 /im HyS9RaFrameworkAgent_epmsystem4b.exe /t /F }
}