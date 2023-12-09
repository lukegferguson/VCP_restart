#Variables
$NGVCP = "C:\NextGen\VCP.exe"
$LogDir = "C:\OSIS"
$LogFile = "VCPRegLog.txt"
$LogPath = "$LogDir\$LogFile"
$NGDirectory = "C:\OSIS\TEST"

#Intro Paragraph
Write-Host "Beginning VCP Registration Script"

#Create directory for logs if not already existing
Write-Host "Verifying log location"
If ((Test-Path -Path $logdir) -eq $FALSE) {
    New-Item -Path 'C:\OSIS' -ItemType Directory | Out-Null
    Write-Host "Created $logdir"
} else {
    Write-Host "Location $logdir already exists"
}

#Initialize Logging
Out-File -FilePath $LogPath -Append -InputObject "$(get-date) ####### Initialize VCP Registraion Script Logging ######"

# VCP Registration process attempts to verify files exist, however it ony checks for the ZIPs, not the unpacked files
#Deleting the ZIP files prior to running VCP registration forces it to re-download the ZIPs and unpack their contents, hopefully correcting anything missing/corrupt
$ZIPs = Get-ChildItem $NGDirectory | Where-Object -Property Name -like *.zip
Out-File -FilePath $LogPath -Append -InputObject "$(get-date) .ZIP files contained in ($NGDirectory): $ZIPs"
$ZIPs | Remove-Item
Out-File -FilePath $LogPath -Append -InputObject "Operation to delete ZIPs complete"

#log .NET Installed before changes, as it may be changed during VCP registration
$DotNetPre = dotnet --list-runtimes
Out-File -FilePath $LogPath -Append -InputObject "$(get-date) .NET Runtimes installed: $DotNetPre"

#Switch user mode to allow for software install, changes how Windows handles INI files
change user /install | Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) $_"




