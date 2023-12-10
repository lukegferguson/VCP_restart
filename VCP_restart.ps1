#Variables
$LogDir = "C:\OSIS"
$LogFile = "VCPRegLog.txt"
$LogPath = "$LogDir\$LogFile"
$VCPScript = "$LogDir\Scripts\VCP_Restart.ps1"
$NGDirectory = "C:\NextGen"
$VCPexe = "C:\NextGen\VCP.exe"
$ScriptVer = "0.5"

#Intro Paragraph
Write-Host "Beginning VCP Registration Script version $ScriptVer"

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
if ($null -eq (Get-ItemProperty $RunOnce -Name "VCPRegProg")){

    #Set Progress tracker at 0
    New-ItemProperty -Path $RunOnce -Name "VCPRegProg" -PropertyType Dword -value 0 | Out-Null
    
    #Write Logs
    Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) Progress registry entry created" 
    Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) VCP Registration Script Progress: $((Get-ItemProperty $RunOnce -Name VCPRegProg).VCPRegProg)"

} else {
    Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) VCP Registration Script Progress: $((Get-ItemProperty $RunOnce -Name VCPRegProg).VCPRegProg)"
}

#Check VCP Registration Progress
$Progress = (Get-ItemProperty $RunOnce -Name "VCPRegProg").VCPRegProg

#Execute this section on first and second script run
#I CANNOT get the if statement to match (0 -or 1) with the same scriptblock, so just copied an entire new elseif for [1] despite being the same code
#[2] Also has a lot of functionality overlap, when I get time, keep it DRY
if ($Progress -eq 0){
        # VCP Registration process attempts to verify files exist, however it ony checks for the ZIPs, not the unpacked files
        #Deleting the ZIP files prior to running VCP registration forces it to re-download the ZIPs and unpack their contents, hopefully correcting anything missing/corrupt
        $ZIPs = Get-ChildItem $NGDirectory | Where-Object -Property Name -like *.zip
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) .ZIP files contained in ($NGDirectory): $ZIPs"
        $ZIPs | Remove-Item
        Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) Operation to delete ZIPs complete"

        #log .NET Installed before changes, as it may be changed during VCP registration
        $DotNet = dotnet --list-runtimes
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) .NET Runtimes currently installed: $DotNet"

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

        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) VCP /r Process Exited"
        
        #Add Script Location to RunOnce to execute again after reboot
        New-ItemProperty -Path $RunOnce -Name "VCPRegScript" -PropertyType String -Value "Powershell.exe $VCPScript"
        Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) Script location added to registry RunOnce: ($VCPScript)" 

        #Increase progress count
        $Progress += 1
        Set-ItemProperty -Path $RunOnce -Name "VCPRegProg" -Value $($progress)
        Out-File -FilePath $LogPath -Append -InputObject "Script will resume when an administrator logs in"
        Out-File -FilePath $LogPath -Append -InputObject "Updated registry to $progress, rebooting..."
        Write-Host "Machine Will REBOOT IN 15 SECONDS, script will resume when an administrator logs in" -ForegroundColor Red
        Start-Sleep -Seconds 15
        Restart-Computer -Force
    }
    elseif ($Progress -eq 1){
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

        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) VCP /r Process Exited"
        
        #Add Script Location to RunOnce to execute again after reboot
        New-ItemProperty -Path $RunOnce -Name "VCPRegScript" -PropertyType String -Value "Powershell.exe $VCPScript"
        Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) Script location added to registry RunOnce: ($VCPScript)" 

        #Increase progress count
        $Progress += 1
        Set-ItemProperty -Path $RunOnce -Name "VCPRegProg" -Value $($progress) | Out-Null
        Out-File -FilePath $LogPath -Append -InputObject "Script will resume when an administrator logs in"
        Out-File -FilePath $LogPath -Append -InputObject "Updated registry to $progress, rebooting..."
        Write-Host "Machine Will REBOOT IN 15 SECONDS, script will resume when an administrator logs in" -ForegroundColor Red
        Start-Sleep -Seconds 15
        Restart-Computer -Force
    }
    elseif ($progress -eq 2) {
        
        #VCP process start (without /r switch)
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) Starting $VCPexe"
        Start-Process -FilePath $VCPexe

        do { Start-Sleep -Seconds 1} while ($null -ne $(get-process VCP*))

        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) VCP Process Exited"

        1..3 | foreach { 
            #Stop any NextGen processes before running again
            $ProcToStop = get-process NG*
            Out-File -FilePath $LogPath -Append -InputObject "$(get-date) NextGen processes running: $($ProcToStop.Name)"
            foreach ($Proc in $ProcToStop){
                Stop-Process -Name $proc.Name
                Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) Stopped process $($Proc.name)"
            

            #VCP process start (without /r switch)
            Out-File -FilePath $LogPath -Append -InputObject "$(get-date) Starting $VCPexe"
            Start-Process -FilePath $VCPexe

            do { Start-Sleep -Seconds 1} while ($null -ne $(get-process VCP*))
            Out-File -FilePath $LogPath -Append -InputObject "$(get-date) VCP Process Exited"}
        
        }

        #log .NET Installed before changes, as it may be changed during VCP registration
        $DotNet = dotnet --list-runtimes
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) .NET Runtimes currently installed: $DotNet"
        
        #Add Script Location to RunOnce to execute again after reboot
        New-ItemProperty -Path $RunOnce -Name "VCPRegScript" -PropertyType String -Value "Powershell.exe $VCPScript"
        Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) Script location added to registry RunOnce: $VCPScript"

        #Increase progress count
        $Progress += 1
        Set-ItemProperty -Path $RunOnce -Name "VCPRegProg" -Value $($progress) | Out-Null
        Out-File -FilePath $LogPath -Append -InputObject "Script will resume when an administrator logs in"
        Out-File -FilePath $LogPath -Append -InputObject "Updated registry to $progress, rebooting..."
        Write-Host "Machine Will REBOOT IN 15 SECONDS, script will resume when an administrator logs in" -ForegroundColor Red
        Start-Sleep -Seconds 15
        Restart-Computer -Force
    }
    
    elseif ($progress -eq 3){
        
        #log .NET Installed before changes, as it may be changed during VCP registration
        $DotNet = dotnet --list-runtimes
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) .NET Runtimes currently installed: $DotNet"
        
        #Switch user mode to allow for software install, changes how Windows handles INI files
        $execute = change user /execute
        Out-File -FilePath $LogPath -Append -InputObject "$(Get-Date) $execute"
        
        #Delete progress tracking registry entry
        Remove-ItemProperty -Path $RunOnce -Name "VCPRegProg" | Out-Null
        Out-File -FilePath $LogPath -Append -InputObject "Removed registry for progress tracking"
        
        #Finalize Logs
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) ##### VCP Registration Complete #####"
    }
    
    else {
        Out-File -FilePath $LogPath -Append -InputObject "$(get-date) real big oops there 😣"
    }


