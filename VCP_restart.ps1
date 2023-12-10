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

#Initialize Logging function 
function Write-Log {
    param ($InputObject)

    Out-File -FilePath $LogPath -Append -InputObject "$(Get-date) $InputObject"
}

Write-Log -InputObject "####### Initialize VCP Registraion Script Logging ######"

#Check Registry for progress
$RunOnce = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
if ($null -eq (Get-ItemProperty -ErrorAction SilentlyContinue $RunOnce -Name "VCPRegProg")){

    #Set Progress tracker at 0
    New-ItemProperty -Path $RunOnce -Name "VCPRegProg" -PropertyType Dword -value 0 | Out-Null
    
    #Write Logs
    Write-Log -InputObject "Progress registry entry created" 
    Write-Log -InputObject "VCP Registration Script Progress: $((Get-ItemProperty $RunOnce -Name VCPRegProg).VCPRegProg)"

} else {
    Write-Log -InputObject "VCP Registration Script Progress: $((Get-ItemProperty $RunOnce -Name VCPRegProg).VCPRegProg)"
}

#Check VCP Registration Progress
$Progress = (Get-ItemProperty $RunOnce -Name "VCPRegProg").VCPRegProg

#Execute this section on first and second script run
#Lot of functionality overlap between 0/1 and 2 loops, when I get time, keep it DRY
switch ($progress){
    {($_ -le 1)} {
         # VCP Registration process attempts to verify files exist, however it ony checks for the ZIPs, not the unpacked files
        #Deleting the ZIP files prior to running VCP registration forces it to re-download the ZIPs and unpack their contents, hopefully correcting anything missing/corrupt
        $ZIPs = Get-ChildItem $NGDirectory | Where-Object -Property Name -like *.zip
        Write-Log -InputObject ".ZIP files contained in ($NGDirectory): $ZIPs"
        $ZIPs | Remove-Item
        Write-Log -InputObject "Operation to delete ZIPs complete"

        #log .NET Installed before changes, as it may be changed during VCP registration
        $DotNet = dotnet --list-runtimes
        Write-Log -InputObject ".NET Runtimes currently installed: $DotNet"

        #Stop any NextGen processes before proceeding
        $ProcToStop = get-process NG*
        Write-Log -InputObject "NextGen processes running: $($ProcToStop.Name)"
        foreach ($Proc in $ProcToStop){
            Stop-Process -Name $proc.Name
            Write-Log -InputObject "Stopped process $($Proc.name)"
        }

        #Switch user mode to allow for software install, changes how Windows handles INI files
        $install = change user /install 
        Write-Log -InputObject "$install"


        #VCP Registration process start
        Write-Log -InputObject "Starting $VCPexe /r"
        Start-Process -FilePath $VCPexe -ArgumentList '/r'

        do { Start-Sleep -Seconds 1} while ($null -ne $(get-process VCP*))

        Write-Log -InputObject "VCP /r Process Exited"
        
        #Add Script Location to RunOnce to execute again after reboot
        New-ItemProperty -Path $RunOnce -Name "VCPRegScript" -PropertyType String -Value "Powershell.exe $VCPScript"
        Write-Log -InputObject "Script location added to registry RunOnce: ($VCPScript)" 

        #Increase progress count
        $Progress += 1
        Set-ItemProperty -Path $RunOnce -Name "VCPRegProg" -Value $($progress)
        Out-File -FilePath $LogPath -Append -InputObject "Script will resume when an administrator logs in"
        Out-File -FilePath $LogPath -Append -InputObject "Updated registry to $progress, rebooting..."
        Write-Host "Machine Will REBOOT IN 15 SECONDS, script will resume when an administrator logs in" -ForegroundColor Red
        Start-Sleep -Seconds 15
        Restart-Computer -Force
    }

    #Run on third iteration of script
    (2) {
         #VCP process start (without /r switch)
         Write-Log -InputObject "Starting $VCPexe"
         Start-Process -FilePath $VCPexe
 
         do { Start-Sleep -Seconds 1} while ($null -ne $(get-process VCP*))
 
         Write-Log -InputObject "VCP Process Exited"
 
         1..3 | foreach {
             
            #Stop any NextGen processes before running again
             $ProcToStop = get-process NG*
             Write-Log -InputObject "NextGen processes running: $($ProcToStop.Name)"
             foreach ($Proc in $ProcToStop){
                 Stop-Process -Name $proc.Name
                 Write-Log -InputObject "Stopped process $($Proc.name)"
 
             #VCP process start (without /r switch)
             Write-Log -InputObject "Starting $VCPexe"
             Start-Process -FilePath $VCPexe
 
             #Wwait for VCP to close before exiting
             do { Start-Sleep -Seconds 1} while ($null -ne $(get-process VCP*))
             Write-Log -InputObject "VCP Process Exited"}
            }
    }
        
    #Fourth iteration of script
    (3) {
            #log .NET Installed before changes, as it may be changed during VCP registration
    $DotNet = dotnet --list-runtimes
    Write-Log -InputObject ".NET Runtimes currently installed: $DotNet"
    
    #Switch user mode to allow for software install, changes how Windows handles INI files
    $execute = change user /execute
    Write-Log -InputObject "$execute"
    
    #Delete progress tracking registry entry
    Remove-ItemProperty -Path $RunOnce -Name "VCPRegProg" | Out-Null
    Out-File -FilePath $LogPath -Append -InputObject "Removed registry for progress tracking"
    
    #Finalize Logs
    Write-Log -InputObject "##### VCP Registration Complete #####"
    }
        
    default {
        Write-Log -InputObject "real big oops there 😣"
    }
}


