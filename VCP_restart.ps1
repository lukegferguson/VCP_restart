#Variables
$LogDir = "C:\OSIS"
$LogFile = "VCPRegLog.txt"
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

#Check Registry for progress
$RunOnce = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
if ( $null -eq (Get-ItemProperty $RunOnce -Name "VCPRegProg")){
    
    #remove any past script references
    Remove-Item -Path  $RunOnce -Name "VCPRegScript" | Out-Null
    
    #Add current script to RunOnce so it will be run on reboot
    #This launches in Notepad currently
    New-ItemProperty -Path $RunOnce -Name "VCPRegScript" -PropertyType String -Value "$PSCommandPath" | out-null
    
    #Set Progress tracker at 0
    New-ItemProperty -Path $RunOnce -Name "VCPRegProg" -PropertyType Dword -value 0 | Out-Null
    
    #Write Logs
    Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) Registry Entries Created" 
    Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) VCP Registration Script Progress: $((Get-ItemProperty $RunOnce -Name VCPRegProg).VCPRegProg)"

} else {
    Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) VCP Registration Script Progress: $((Get-ItemProperty $RunOnce -Name VCPRegProg).VCPRegProg)"
}

$Progress = (Get-ItemProperty $RunOnce -Name "VCPRegProg").VCPRegProg

#Execute this section on first and second script run
if ($progress -eq 0 -or 1){
        # VCP Registration process attempts to verify files exist, however it ony checks for the ZIPs, not the unpacked files
        #Deleting the ZIP files prior to running VCP registration forces it to re-download the ZIPs and unpack their contents, hopefully correcting anything missing/corrupt
        $ZIPs = Get-ChildItem $NGDirectory | Where-Object -Property Name -like *.zip
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) .ZIP files contained in ($NGDirectory): $ZIPs"
        $ZIPs | Remove-Item
        Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) Operation to delete ZIPs complete"

        #log .NET Installed before changes, as it may be changed during VCP registration
        $DotNet = dotnet --list-runtimes
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) .NET Runtimes installed: $DotNet"

        #Stop any NextGen processes before proceeding
        $ProcToStop = get-process NG*
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) NextGen processes running: $($ProcToStop.Name)"
        foreach ($Proc in $ProcToStop){
            Stop-Process -Name $proc.Name
            Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) Stopped process $($Proc.name)"
        }

        #Switch user mode to allow for software install, changes how Windows handles INI files
        $install = change user /install 
        Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) $install"


        #VCP Registration process start
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) Starting $VCPexe /r"
        Start-Process -FilePath $VCPexe -ArgumentList '/r'

        do { Start-Sleep -Seconds 1} while ($null -ne $(get-process VCP*))

        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) VCP Process Exited"

        #Increase progress count
        $Progress += 1
        Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "VCPRegProg" -Value $($progress)
        Out-File -FilePath $LogPath -Append -InputObject "Updated registry to $progress, rebooting..."
        Out-File -FilePath $LogPath -Append -InputObject "Script will resume when an administrator logs in"
        Write-Host "Machine Will REBOOT IN 15 SECONDS, script will resume when an administrator logs in" -ForegroundColor Red
        Start-Sleep -Seconds 15
        #Restart-Computer -Force
    }
    
    elseif ( $progress -eq 2) {
        #TODO vcp.exe without switch then rebooting again
        #Update reg switch
        #timing or repeats??
    }
    
    elseif ($progress -eq 3){
        #TODO Delete regi switches and run set user /Execute
        #finalize logs
    }
    
    else {
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) real big oops there"
    }


