#Variables
$LogDir = "C:\OSIS"
$LogFile = "$(get-date -Format "yyyy.MM.dd").VCPRegLog.txt"
$LogPath = "$LogDir\$LogFile"
$NGDirectory = "C:\NextGen"
$VCPexe = "C:\NextGen\VCP.exe"

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
$DotNet = dotnet --list-runtimes
Out-File -FilePath $LogPath -Append -InputObject "$(get-date) .NET Runtimes installed: $DotNet"

#Stop any NextGen or VCP processes before proceeding
$ProcToStop = get-process NG*
Out-File -FilePath $LogPath -Append -InputObject "$(get-date) NextGen processes running: $($ProcToStop.Name)"
foreach ($Proc in $ProcToStop){
    Stop-Process -Name $proc.Name
    Out-File -FilePath $LogPath -Append -InputObject "Stopped process $($Proc.name)"
}
#Out-File -FilePath $LogPath -Append -InputObject "$(get-date) Stopped processes $StoppedProc"

#Switch user mode to allow for software install, changes how Windows handles INI files
$install = change user /install 
Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) $install"


#VCP Registration process start
Out-File -FilePath $LogPath -Append -InputObject "$(get-date) Starting $VCPexe /r"
$VCPProc = Start-Process -wait -NoNewWindow -FilePath $VCPexe -ArgumentList '/r' -PassThru
Out-File -FilePath $LogPath -Append -InputObject "$(get-date) Started $($VcpProc.name)"
